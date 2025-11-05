Attribute VB_Name = "SubRutinas"
Option Explicit

' Teclas Especiales del Teclado
' SystemDirectory
' ShiftDown
' AltDown
' CtrlDown

'Letras Impresora
'*** TIPOS DE LETRAS IMPRESORA
'Print #1, Chr(27) & Chr(120) & Chr(0) 'Draft
'Print #1, Chr(27) & Chr(77)           '12 CPI
'Print #1, Chr(27) & Chr(80)           '10 CPI
'Print #1, Chr(27) & Chr(15)           'Comprimido
'Print #1, Chr(27) & Chr(14)           'Ancho Double
'Print #1, Chr(27) & Chr(69)           'Negrita
'Print #1, Chr(14)                     'Agrandar

'Resetear Impresiones
'Chr(27) & Chr(18)                     'Cancela Comprimido
'Chr(27) & Chr(20)                     'Cancela Ancho double
'Chr(27) & Chr(70)                     'Cancela negrita
'Chr(18)                               'Cancela Agrandar

'Estos caracteres se los debe sumar a la cadena que deseas imprimir y al final
'de la cadena los caracteres para cancelar...

'Progreso_Barra.Mensaje_Box = "Consultando el Balance"
'Progreso_Iniciar
'Progreso_Esperar
'Progreso_Final
 
''     SetAdoAddNew "Nombre_Tabla"
''     SetAdoFields "Campo", Valor
''     SetAdoFields "CodigoU", CodigoUsuario
''     SetAdoFields "Periodo", Periodo_Contable
''     SetAdoFields "Item", NumEmpresa
''     SetAdoUpdate
 
'      & "InStr 'MIOS': '" & InStr(Cadena, "MIOS") & "'"
 
'       'Elimina la actualizacion anterior si hay conexion
'        Eliminar_Si_Existe_File RutaSistema & "\BASES\UPDATE_DB\*.*"
'
'       'Mostramos en el label el path del directorio actual donde estamos ubicados en el servidor
'        Progreso_Barra.Mensaje_Box = .GetDirectorioActual
'       .Mostar_Estado_FTP ProgressBarEstado, LstStatud
'       'Le indicamos el ListView donde se listarán los archivos
'        Set .ListView = LstVwFTP
'        Progreso_Barra.Mensaje_Box = "Buscando directorio en el servidor"
'       .Mostar_Estado_FTP ProgressBarEstado, LstStatud
'
'       '-------------------------------------------------------
'       'Esta opcion solo baja la actualizacion del servidor erp
'       '=======================================================
'        Progreso_Barra.Mensaje_Box = "Eliminando Version anterior"
'       .Mostar_Estado_FTP ProgressBarEstado, LstStatud
'       .CambiarDirectorio "/SISTEMA/BASES/UPDATE_DB/"
'       .ListarArchivos
'        For I = 1 To LstVwFTP.ListItems.Count
'            TextoFile = ""
'            FileOrigen = LstVwFTP.ListItems(I)
'            FileDestino = RutaSistema & "\BASES\UPDATE_DB\" & LstVwFTP.ListItems(I)
'            Extension = RightStrg(FileOrigen, 3)
'            Select Case Extension
'              Case "dbs", "txt", "upd", "sql"
'                   Progreso_Barra.Mensaje_Box = "Descargando: " & FileOrigen
'                  .Mostar_Estado_FTP ProgressBarEstado, LstStatud
'                  .ObtenerArchivo FileOrigen, FileDestino, True
'            End Select
'        Next I
'       .Desconectar

Public Sub screen_size()
    MDI_X_Max = MDIFormulario.ProgressBarEstado.width - 100
    MDI_Y_Max = Screen.Height - 1940
End Sub

Public Function Color_Dias_Restantes(vDiasRestantes As Long) As Long
Dim vColor As Long
    Select Case vDiasRestantes
      Case Is <= 10: vColor = QBColor(12)
      Case 11 To 30: vColor = Fucsia_Claro
      Case 31 To 60: vColor = QBColor(14)
      Case 61 To 90: vColor = QBColor(6)
      Case 91 To 120: vColor = QBColor(9)
      Case 121 To 150: vColor = QBColor(10)
      Case Is >= 151: vColor = QBColor(2)
    End Select
    Color_Dias_Restantes = vColor
End Function

'''Public Function Solo_Letras_Numeros(ByVal KeyAscii) As Integer
'''   If KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 11 Or KeyAscii >= 32 Then
'''      Solo_Letras_Numeros = KeyAscii
'''   Else
'''      Solo_Letras_Numeros = 0
'''   End If
'''   '
'''End Function

Public Function TrimStrg(Cadena As Variant) As String
Dim Resultado As String
    Resultado = ""
    If Len(Cadena) > 0 Then Resultado = Trim$(Cadena)
    TrimStrg = Resultado
End Function

Public Function LeftStrg(Cadena As Variant, CantStr As Long) As String
Dim Resultado As String
    Resultado = ""
    If CantStr <= Len(Cadena) Then Resultado = Trim$(Left$(Cadena, CantStr))
    LeftStrg = Resultado
End Function

Public Function RightStrg(Cadena As Variant, CantStr As Long) As String
Dim Resultado As String
    Resultado = ""
    If CantStr <= Len(Cadena) Then Resultado = Trim$(Right$(Cadena, CantStr))
    RightStrg = Resultado
End Function

Public Function MidStrg(Cadena As Variant, InicioStr As Long, Optional CantStr As Long) As String
Dim Resultado As String
    If Len(Cadena) > 0 And CantStr > 0 Then
       If InicioStr > 0 Then Resultado = Mid$(Cadena, InicioStr, CantStr) Else Resultado = Mid$(Cadena, 1, CantStr)
    Else
       Resultado = ""
    End If
    MidStrg = Resultado
End Function

Public Function UCaseStrg(Cadena As String) As String
Dim Resultado As String
    If Len(Cadena) > 0 Then Resultado = Trim$(UCase$(Cadena)) Else Resultado = ""
    UCaseStrg = Resultado
End Function

Public Function LCaseStrg(Cadena As String) As String
Dim Resultado As String
    If Len(Cadena) > 0 Then Resultado = Trim$(LCase$(Cadena)) Else Resultado = ""
    LCaseStrg = Resultado
End Function

Public Function ULCase(TextoConversion As String) As String
Dim TI As Long
Dim Mayusc As Boolean
Dim CadAux As String
Dim Caracter As String
Dim TextoULCase As String
  TextoULCase = LCase$(TextoConversion)
  If TextoULCase = "" Then TextoULCase = Ninguno
  CadAux = ""
  Mayusc = True
  For TI = 1 To Len(TextoULCase)
      Caracter = MidStrg(TextoULCase, TI, 1)
      If Mayusc Then
         Caracter = UCaseStrg(Caracter)
         Mayusc = False
      End If
      CadAux = CadAux & Caracter
      If Caracter = " " Then Mayusc = True
  Next TI
  ULCase = CadAux
End Function

Public Function Bloquear_Control() As Boolean
Dim OpcionBloq As Boolean
    OpcionBloq = False
    Select Case EstadoEmpresa
      Case "BLOQ", "ONLY": OpcionBloq = True
    End Select
    Bloquear_Control = OpcionBloq
End Function

Public Function EsUnEmail(NEmail As String) As Boolean
Dim esMail As Boolean

  esMail = False
  If Len(NEmail) > 3 Then
     For I = 1 To Len(NEmail)
         If MidStrg(NEmail, I, 1) = "@" Then
            esMail = True
            Exit For
         End If
     Next I
     Select Case MidStrg(NEmail, Len(NEmail))
       Case ".", ",": esMail = False
     End Select
  End If
  EsUnEmail = esMail
End Function

Public Sub Insertar_Mail(ListaMails As String, InsertarMail As String)

   If (InStr(ListaMails, InsertarMail) = 0) And (InStr(InsertarMail, "@") > 0) And (Len(Trim(InsertarMail)) > 3) Then
      ListaMails = ListaMails & InsertarMail & ";"
   Else
      TMail.ListaError = TMail.ListaError & TMail.Destinatario & ": " & InsertarMail & vbCrLf
   End If
End Sub

Public Sub Insertar_Cadena(ListaMails As String, InsertarMail As String)
   If (InStr(ListaMails, InsertarMail) = 0) And (Len(InsertarMail) > 3) Then ListaMails = ListaMails & InsertarMail & "/"
End Sub

Public Sub Master_Documento_PDF(Optional PDF_Nombre_Documento As String, _
                                Optional PDF_Titulo As String, _
                                Optional PDF_TipoDeLetra As String, _
                                Optional PDF_VerDocumento As Boolean)
   RatonReloj
   Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
   Titulo = "IMPRESION"
   Bandera = False
   SetPrinters.Show 1
   If PonImpresoraDefecto(SetNombrePRN) Then
      RatonReloj
     'Generamos el documento
     'tPrint.TipoImpresion = Es_PDF
      tPrint.NombreArchivo = PDF_Nombre_Documento
      tPrint.TituloArchivo = PDF_Titulo
      tPrint.TipoLetra = PDF_TipoDeLetra
      tPrint.OrientacionPagina = Orientacion_Pagina
      tPrint.PaginaA4 = True
      tPrint.EsCampoCorto = False
      tPrint.VerDocumento = PDF_VerDocumento
      Set cPrint = New cImpresion
      cPrint.iniciaImpresion
     
      cPrint.printImagen LogoTipo, 1, 1, 5, 2
      PosLinea = 3.5
      cPrint.colorDeLetra = Negro
      cPrint.tipoNegrilla = True
      If cPrint.anchoTexto(UCaseStrg(RazonSocial)) > 9 Then
         cPrint.letraTipo PDF_TipoDeLetra, 6
         cPrint.printTexto 1.1, PosLinea, UCaseStrg(RazonSocial)
         PosLinea = PosLinea + 0.35
      Else
         cPrint.printTexto 1.1, PosLinea, UCaseStrg(RazonSocial)
         PosLinea = PosLinea + 0.35
      End If
      cPrint.generarBarras "1234567890", cC128_A, 10, CDbl(PosLinea), 10, 0.8
      cPrint.printLinea 1, 0.1, 20.5, 28, Negro, "B"
      PosLinea = PosLinea + 0.25
      cPrint.paginaNueva
      cPrint.finalizaImpresion
      RatonNormal
      MsgBox "Proceso terminado"
   End If
End Sub

Public Sub Master_Progreso_Barras()
  Progreso_Iniciar
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = 100
  
  Progreso_Barra.Mensaje_Box = "Mensaje ejemplo"
  Progreso_Esperar
  
  Progreso_Final
End Sub

'''Public Sub Ctrl_C_Grid(DGCopy As DataGrid)
'''Dim IFields As Integer
''' RatonReloj
''' ReDim CopyGrid(DGCopy.columnCount) As Campos_Tabla
''' For IFields = 0 To DGCopy.columnCount - 1
'''     CopyGrid(IFields).Valor = DGCopy.Columns(IFields)
'''     CopyGrid(IFields).Campo = DGCopy.Columns(IFields).DataField
''' Next IFields
''' RatonNormal
'''End Sub

Public Sub Ctrl_V_Grid(Nombre_Tabla As String)
Dim IFields As Integer
  RatonReloj
  If UBound(CopyGrid) > 0 Then
     SetAdoAddNew Nombre_Tabla
     For IFields = 0 To UBound(CopyGrid) - 1
         If IFields = (UBound(CopyGrid) - 1) And CopyGrid(IFields).Campo <> "Item" Then CopyGrid(IFields).Valor = 0
         SetAdoFields CopyGrid(IFields).Campo, CopyGrid(IFields).Valor
     Next IFields
     SetAdoUpdate
  End If
  RatonNormal
End Sub

Public Sub Grabar_Consulta_Archivo(Archivo As String, _
                                   Consulta_SQL As String)
Dim NumFile As Integer
Dim RutaGeneraFile As String
RatonReloj
RutaGeneraFile = RutaSysBases & "\TEMP\" & Archivo & ".sql"
'MsgBox RutaGeneraFile
NumFile = FreeFile
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
     Print #NumFile, Consulta_SQL;
     Print #NumFile, ";" & vbCrLf
Close #NumFile
RatonNormal
End Sub

Public Function Leer_Campo_Empresa(Campo_de_Busqueda As String) As Variant
Dim AdoRegs As ADODB.Recordset
Dim Strgs As String
Dim Valor_Retorno As Variant
 Valor_Retorno = Null
'MsgBox Campo_de_Busqueda & vbCrLf & NumEmpresa
 If Len(NumEmpresa) > 1 Then
    If Campo_de_Busqueda <> "" And Val(NumEmpresa) > 0 Then
       Strgs = "SELECT " & Campo_de_Busqueda & " " _
             & "FROM Empresas " _
             & "WHERE Item = '" & NumEmpresa & "' "
       Select_AdoDB AdoRegs, Strgs
       If AdoRegs.RecordCount > 0 Then Valor_Retorno = AdoRegs.fields(Campo_de_Busqueda)
       AdoRegs.Close
    End If
 End If
 If IsNull(Valor_Retorno) Then Valor_Retorno = Ninguno
 Leer_Campo_Empresa = Valor_Retorno
End Function

Public Function Leer_Campo_Educativo(Campo_de_Busqueda As String) As Variant
Dim Strgs As String
Dim Valor_Retorno As Variant
Dim AdoRegs As ADODB.Recordset
 Valor_Retorno = Null
 If Campo_de_Busqueda <> "" And Val(NumEmpresa) > 0 Then
   'MsgBox Campo_de_Busqueda
    Set AdoRegs = New ADODB.Recordset
    AdoRegs.CursorType = adOpenStatic
    AdoRegs.CursorLocation = adUseClient
    Strgs = "SELECT * " _
          & "FROM Catalogo_Periodo_Lectivo " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
    AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
    If AdoRegs.RecordCount > 0 Then
       Valor_Retorno = AdoRegs.fields(Campo_de_Busqueda)
    End If
    AdoRegs.Close
 End If
 Leer_Campo_Educativo = Valor_Retorno
End Function

Public Function Leer_Campo_Cliente(Campo_de_Busqueda As String, Patron_Busqueda As String) As String
Dim Strgs As String
Dim Valor_Retorno  As String
Dim AdoRegs As ADODB.Recordset
 RatonReloj
 Valor_Retorno = ""
 If Len(Campo_de_Busqueda) > 1 And Len(Patron_Busqueda) > 1 Then
  'MsgBox Campo_de_Busqueda
   Set AdoRegs = New ADODB.Recordset
   AdoRegs.CursorType = adOpenStatic
   AdoRegs.CursorLocation = adUseClient
   Strgs = "SELECT Codigo, Cliente,CI_RUC " _
         & "FROM Clientes " _
         & "WHERE " & Campo_de_Busqueda & " = '" & Patron_Busqueda & "' "
   AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
   If AdoRegs.RecordCount > 0 Then
      Valor_Retorno = AdoRegs.fields("Cliente")
      Codigos = AdoRegs.fields("Codigo")
      CodigoP = AdoRegs.fields("CI_RUC")
   End If
   AdoRegs.Close
 End If
 RatonNormal
 Leer_Campo_Cliente = Valor_Retorno
End Function

Public Sub Fecha_Procesos(Detalle_Fecha As String, _
                          Fecha_Inicial As String, _
                          Fecha_Final As String)
  RatonReloj
  sSQL = "DELETE * " _
       & "FROM Fechas_Balance " _
       & "WHERE Detalle = '" & Detalle_Fecha & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP sSQL
  sSQL = "INSERT INTO Fechas_Balance (Periodo,Item,Detalle,Fecha_Inicial,Fecha_Final,Cerrado) " _
       & "VALUES " _
       & "('" & Periodo_Contable & "','" & NumEmpresa & "','" & Detalle_Fecha & "'," _
       & "#" & BuscarFecha(Fecha_Inicial) & "#,#" & BuscarFecha(Fecha_Final) & "#,0) "
  Ejecutar_SQL_SP sSQL
  RatonNormal
End Sub

Public Sub Imprimir_Lineas_Campos(Linea As Single)
Dim CInicio As Integer
Dim AltoLetras As Single
  AltoLetras = Printer.TextHeight("H")
  If AltoLetras <= 0 Then AltoLetras = 0.35
  For CInicio = 0 To CantCampos
      Printer.Line (Ancho(CInicio), Linea - 0.05)-(Ancho(CInicio), Linea + AltoLetras + 0.05), Negro
  Next CInicio
End Sub

''     SetAdoAddNew "Trans_Entrada_Salida"
''     SetAdoFields "ES", UCaseStrg(TipoTrans)
''     SetAdoFields "Codigo", Modulos
''     SetAdoFields "Hora", Format$(Time, FormatoTimes)
''     SetAdoFields "Fecha", FechaSistema
''     If Proceso <> "" Then SetAdoFields "Proceso", TrimStrg(MidStrg(Proceso, 1, 60))
''     If Tarea = Ninguno Then
''        SetAdoFields "Tarea", "Inicio de Sección"
''     Else
''        SetAdoFields "Tarea", TrimStrg(MidStrg(Tarea, 1, 60))
''     End If
''     SetAdoFields "Credito_No", TMail.Credito_No
''     SetAdoFields "CodigoU", CodigoUsuario
''     SetAdoFields "Periodo", Periodo_Contable
''     SetAdoFields "Item", NumEmpresa
''     SetAdoUpdate
    'MsgBox strServidor

Public Sub Control_Procesos(TipoTrans As String, Proceso As String, Optional Tarea As String, Optional Credito_No As String)
Dim BD1MySQL As ADODB.Connection
Dim AdoSMTP As ADODB.Recordset
Dim Modulos As String
'Dim Mifecha1 As String
'Dim MiHora1 As String
Dim NombreUsuario1 As String
Dim PausaMails As Long
Dim TimeIni As Single

  If NumEmpresa = "" Then NumEmpresa = Ninguno
  If TMail.Credito_No = "" Then TMail.Credito_No = Ninguno
  If Modulo <> Ninguno And TipoTrans <> Ninguno And NumEmpresa <> Ninguno Then
     If Proceso = Ninguno Then Proceso = "Procesando..." Else Proceso = TrimStrg(MidStrg(Proceso, 1, 120))
     Tarea = TrimStrg(MidStrg(Tarea, 1, 120))
     If Tarea = "" Then Tarea = Ninguno
     If Len(Proceso) > 1 Then
        'If Ping_IP(strServidor) Then
           Modulos = Modulo
           NombreUsuario1 = TrimStrg(MidStrg(NombreUsuario, 1, 60))
           TipoTrans = UCaseStrg(TipoTrans)
'           Mifecha1 = Format(Date, "yyyy-mm-dd")
 '          MiHora1 = Format(Time, "hh:mm:ss")
           If Credito_No = "" Then Credito_No = Ninguno
           
           Cadena = IP_PC.Nombre_PC & vbCrLf _
                  & IP_PC.IP_PC & vbCrLf _
                  & IP_PC.Max_IP & vbCrLf _
                  & IP_PC.InterNet & vbCrLf _
                  & IP_PC.MAC_PC & vbCrLf _
                  & String(20, "=") & vbCrLf
'''           For I = 0 To UBound(IP_PC.Lista_IPs)
'''               Cadena = Cadena & IP_PC.Lista_IPs(I) & vbCrLf
'''           Next I
           Control_Procesos_SP_MySQL IP_PC.IP_PC, TipoTrans, Proceso, Tarea, Credito_No

'''           sSQL = "INSERT INTO acceso_pcs (IP_Acceso,CodigoU,Item,Aplicacion,RUC,ES,Tarea,Proceso,Credito_No,Periodo) " _
'''                & "VALUES ('" & IP_PC.IP_PC & "','" & CodigoUsuario & "','" & NumEmpresa & "'," & "'" & Modulo & "','" & RUC _
'''                & "','" & TipoTrans & "','" & Tarea & "','" & Proceso & "','" & Credito_No & "','" & Periodo_Contable & "');"
'''          'MsgBox Cadena & vbCrLf & vbCrLf & "CONSULTA A REALIZAR:" & vbCrLf & sSQL
'''           Conectar_Ado_Execute_MySQL sSQL
'''
'''           If TipoTrans = "EM" Then
'''              PausaMails = 0
'''              sSQL = "SELECT RUC, (COUNT(RUC) % 10) As PausaMails " _
'''                   & "FROM acceso_pcs " _
'''                   & "WHERE Fecha_Hora = '" & BuscarFecha(FechaSistema) & "' " _
'''                   & "AND RUC = '" & RUC & "' " _
'''                   & "AND ES IN ('EM') "
'''              Select_AdoDB_MySQL AdoSMTP, sSQL
'''              If AdoSMTP.RecordCount > 0 Then PausaMails = AdoSMTP.Fields("PausaMails")
'''              AdoSMTP.Close
'''             'Hacer pausa si ya tiene 20 envios
'''              If PausaMails = 0 Then
'''                 RatonReloj
'''                 Sleep 10000
'''                'MsgBox Format(Timer - TimeIni, "mm:ss")
'''                 RatonNormal
'''              End If
'''           End If
        'End If
     End If
  End If
End Sub

Function FormatInt(Valor_Convertir As Variant)
  FormatInt = Format$(Valor_Convertir, "##0")
End Function

Function FormatDbl(Valor_Convertir As Variant)
  FormatDbl = Format$(Valor_Convertir, "#,##0")
End Function

Function FormatDec(Valor_Convertir As Variant, NumDecimales As Byte)
  FormatDec = Format$(Valor_Convertir, "#,##0." & String$(NumDecimales, "0"))
End Function

Public Function Truncar(ByVal Numero As Double, Optional ByVal Decimales As Byte = 0) As Double
Dim lngPotencia As Long
    lngPotencia = 10 ^ Decimales
    Numero = Int(Numero * lngPotencia)
    Truncar = Numero / lngPotencia
End Function

Function Redondear(Valor As Variant, Optional Decim As Byte) As Variant
Dim Valor_Redondeo As Double
  If Decim <= 0 Then Decim = 0
  If Decim >= 6 Then Decim = 6
  Valor_Redondeo = Round(Valor, 6)
  Valor_Redondeo = Val(Format$(Valor_Redondeo, "##0." & String$(Decim, "0")))
 'MsgBox Valor & vbCrLf & Valor_Redondeo
  Redondear = Valor_Redondeo
End Function

Function Redondear_2Dec(Valor As Variant) As Variant
Dim Parte_Entera As Currency
Dim Parte_Decimal As Currency
Dim Valor_Redondeo As Currency
Dim Str_Entero As String
Dim PosDec As Integer
  
  Valor_Redondeo = Round(Valor, 6)
  
  Str_Entero = CStr(Valor_Redondeo)
  PosDec = InStr(Str_Entero, ".")
  Parte_Entera = Int(Valor_Redondeo)
  If PosDec > 0 Then Str_Entero = MidStrg(Str_Entero, PosDec + 1, Len(Str_Entero)) Else Str_Entero = "0"
      If Val(MidStrg(Str_Entero, 3, 1)) > 0 Then
         Str_Entero = MidStrg(Str_Entero, 1, 2)
         Valor_Redondeo = Parte_Entera + CDbl("0." & Str_Entero)
      Else
         Valor_Redondeo = Parte_Entera + CDbl("0." & Str_Entero)
      End If
  Redondear_2Dec = Valor_Redondeo
End Function

Function ExisteFormulario(Nombre As String) As Boolean
Dim F As Form
On Error GoTo ErrorExisteFormulario

ExisteFormulario = False
Cadena = ""
'Recorremops los formularios activos
For Each F In Forms
    Cadena = Cadena & UCaseStrg(F.Name) & vbCrLf
    If UCaseStrg(F.Name) = UCaseStrg(Nombre) Then
        ExisteFormulario = True
        Exit For
    End If
Next

'MsgBox Cadena
Exit Function
ErrorExisteFormulario:
     MsgBox "ExisteFormulario", Err, Error
End Function

Public Sub Imprimir_Texto_Lineas(Texto_Imp As String)
Dim IniX As Single
Dim IniY As Single
Dim Texto As String
Dim LineTexto As String
Dim CampoTexto As String
Dim AnchoDeLinea As Single
On Error GoTo Errorhandler
RatonReloj
SQLMsg1 = "IMPRESION DE ERRORES"
SQLMsg2 = "Fecha de Error del Archivo: " & Mifecha
SQLMsg3 = "Fecha de Impresion: " & FechaSistema & " - Usuario: " & NombreUsuario
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
Escala_Centimetro 1, TipoCourierNew, 8
'Iniciamos la impresion
AnchoDeLinea = 18: Pagina = 1
IniX = 1: IniY = 1
Encabezado_Documento IniX, IniY, 19
LineTexto = Texto_Imp
Texto = LineTexto
J = Len(Texto)
I = 1: K = 1
LineTexto = ""
Do While I < J
   Printer.FontSize = 8
   Printer.FontName = TipoCourier
   Caracter = MidStrg(Texto, I, 1)
   CampoTexto = MidStrg(Texto, I, 3)
   LineTexto = LineTexto & Caracter
   If Printer.TextWidth(LineTexto) > AnchoDeLinea Or Asc(Caracter) = 13 Or Asc(Caracter) = 10 Then
      If Printer.TextWidth(LineTexto) > AnchoDeLinea Then
         K = Len(LineTexto)
         If K > 0 Then
            Do
               K = K - 1
               I = I - 1
            Loop Until K < 2 Or MidStrg(LineTexto, K, 1) = " "
            LineTexto = MidStrg(LineTexto, 1, K)
         End If
      End If
      'MsgBox IniX & vbCrLf & PosLinea & vbCrLf & LineTexto
      Printer.CurrentX = IniX
      Printer.CurrentY = PosLinea
      Printer.Print LineTexto
      PosLinea = PosLinea + Printer.TextHeight("H") + 0.1
      LineTexto = ""
      If Asc(Caracter) = 13 Then I = I + 1
   End If
   If PosLinea >= LimiteAlto Then
      Printer.NewPage
      PosLinea = IniY + 2
      LineTexto = ""
   End If
   I = I + 1
Loop
'Producto = InsertarLinea
Printer.EndDoc
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub ImprimirCodigoBarra(PosXo As Single, _
                               PosYo As Single, _
                               ValorBarra As String, _
                               PictBar As PictureBox)
'Printer.PaintPicture Clipboard.GetData, PosXo + 0.01, PosYo + 0.01
Dim PosSup, PosIzq As Single
  PictBar.Cls
  PictBar.AutoRedraw = True
  PictBar.FontBold = False
  PictBar.Picture = Clipboard.GetData
  PosIzq = ((PictBar.width - PictBar.TextWidth(ValorBarra)) / 2)
  PosSup = PictBar.Height
  PictBar.Height = PictBar.Height * 1.6
  If PosIzq <= 0 Then PosIzq = 0.01
  If PosSup <= 0 Then PosSup = 0.01
  PictBar.ForeColor = 0
  PictBar.CurrentX = PosIzq
  PictBar.CurrentY = PosSup
  PictBar.FontName = TipoArial
  PictBar.Font = 5
  PictBar.Print ValorBarra
  Printer.PaintPicture PictBar.Image, PosXo + 0.01, PosYo + 0.01
End Sub


Public Sub Imprimir_Codigos_De_Activos(AdoRangoCodigo As Adodc, _
                                       Cantidad As Integer, _
                                       PictBar As PictureBox)
Dim SizeLetra As Integer
Dim PosSup, PosIzq, DivxFila As Single
Dim CodigoBarra As String
SizeLetra = 6
On Error GoTo Errorhandler
RatonReloj
If Cantidad < 3 Then Cantidad = 3
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   InicioX = 1: InicioY = 0: Pagina = 1
   Escala_Centimetro 1, TipoTimes, 10
   Printer.FontName = TipoArialNarrow
   If Cantidad < 3 Then Cantidad = 3
   DivxFila = (LimiteAncho - 1) / Cantidad
   PosLinea = 1: PosColumna = 1: Contador = 1
   With AdoRangoCodigo.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           CodigoBarra = .fields("Codigo_Barra")
           CodigoBarra = Replace(CodigoBarra, "Ñ", "N")
           Printer.FontBold = True
          'Elaboración del grafico del codigo de barra
''           Code39.AlturaBarra = 25
''           Code39.TamBarra = 1
''           Code39.ColorCodigo = "N"
''           Code39.ValorCodigo = CodigoBarra
''           Code39.RealizarCodigo
''           PictBar.Cls
''           PictBar.AutoRedraw = True
''           PictBar.FontBold = False
''           PictBar.Picture = Clipboard.GetData
          'Nombre del Producto
           Producto = .fields("Producto")
           Printer.FontSize = 8
           J = Len(Producto)
           Do While (Printer.TextWidth(MidStrg(Producto, 1, J)) > 4.5) And (J > 0)
              J = J - 1
           Loop
           Producto = MidStrg(Producto, 1, J)
          'Dibujamos el la pagina de la impresora
          'Printer.ForeColor = 0
          'Nombre de la Empresa
           Printer.FontSize = 8
           PosSup = 0.1
           'MsgBox "X: " & PosColumna & " Y: " & PosLinea
           If UCaseStrg(Empresa) = UCaseStrg(NombreComercial) Then
              Printer.CurrentX = PosColumna
              Printer.CurrentY = PosLinea
              Printer.Print "(" & NumEmpresa & ") " & Empresa
           Else
              Printer.CurrentX = PosColumna
              Printer.CurrentY = PosLinea
              Printer.Print "(" & NumEmpresa & ") " & Empresa
              If NombreComercial <> Ninguno Then
                 PosSup = PosSup + Printer.TextHeight("H")
                 Printer.CurrentX = PosColumna
                 Printer.CurrentY = PosLinea + PosSup
                 Printer.Print NombreComercial
              End If
           End If
          'Nombre del Producto
           Printer.FontSize = 8
           Printer.CurrentX = PosColumna
           Printer.CurrentY = PosLinea + 1
           Printer.Print Producto
           Printer.CurrentX = PosColumna
           Printer.CurrentY = PosLinea + 1.4
           Printer.Print .fields("Tipo")
           Printer.CurrentX = PosColumna
           Printer.CurrentY = PosLinea + 1.8
           Printer.Print .fields("Ubicacion")
          'Codigo de Barra
           Printer.PaintPicture PictBar.Image, PosColumna, PosLinea + 2.2
          'Codigo de Barra en Texto
           Printer.FontSize = 9
           Printer.CurrentX = PosColumna
           Printer.CurrentY = PosLinea + 2.9
           Printer.Print " Codigo: " & CodigoBarra
           Printer.CurrentX = PosColumna
           Printer.CurrentY = PosLinea + 3.3
           Printer.Print .fields("Nombre_Responsable")
           
           Contador = Contador + 1
           If Contador > Cantidad Then
              PosLinea = PosLinea + 5
              PosColumna = 1
              Contador = 1
           Else
              PosColumna = PosColumna + DivxFila
           End If
           If PosLinea >= LimiteAlto Then
             Printer.NewPage
             Pagina = Pagina + 1
             PosLinea = 1
             Contador = 1
             PosColumna = 1
           End If
          .MoveNext
        Loop
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

Public Sub Imprimir_Codigos_Estanteria(GrupoProducto As String)
Dim DatasC As ADODB.Recordset
Dim SizeLetra As Integer
Dim PosSup, PosIzq As Single
Dim CodigoBarra As String
Dim CantEtiquetaX As Byte
Dim CantEtiquetaY As Byte
Dim EtiquetaAncho As Single
Dim EtiquetaAlto As Single
Dim PDF_Titulo As String
Dim PDF_Nombre_Documento As String

SizeLetra = 6
CantEtiquetaX = 5
CantEtiquetaY = 10
EtiquetaAncho = 3.8
EtiquetaAlto = 2.5

On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
  'TipoHelvetica
   InicioX = 1: InicioY = 0
   RatonReloj
   Progreso_Iniciar
  'Geneeramos el documento
   tPrint.TipoImpresion = Es_Printer
   tPrint.NombreArchivo = "Codigos de Barra"
   tPrint.TituloArchivo = "Estanteria de Codigos de Barra"
   tPrint.OrientacionPagina = Orientacion_Pagina
   tPrint.TipoLetra = TipoTimes
   tPrint.PorteLetra = 8
   tPrint.VerDocumento = True
   tPrint.PaginaA4 = True
   tPrint.EsCampoCorto = True
   
   Set cPrint = New cImpresion
   cPrint.iniciaImpresion
   
   sSQL = "SELECT * " _
        & "FROM Catalogo_Productos " _
        & "WHERE TC = 'P' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND LEN(Codigo_Barra) > 1 " _
        & "AND MidStrg(Codigo_Inv,1," & Len(GrupoProducto) & ")='" & GrupoProducto & "' " _
        & "ORDER BY Codigo_Inv "
  'MsgBox sSQL
   Select_AdoDB DatasC, sSQL
   If DatasC.RecordCount > 0 Then
      Progreso_Barra.Valor_Maximo = DatasC.RecordCount
      cPrint.tipoNegrilla = True
      cPrint.colorDeLetra = Negro
      Contador = 1
      Pagina = 1
      PosLinea = 1
      Contador = 1
      PosColumna = 1
      DatasC.MoveFirst
      I = 1: J = 1
      PosSup = 1.2
      PosIzq = 0.6
      Do While Not DatasC.EOF
         CodigoBarra = DatasC.fields("Codigo_Barra")
         Precio = DatasC.fields("PVP")
         If DatasC.fields("IVA") Then Precio = Precio + (Precio * Porc_IVA)
         CodigoP = "PVP $ " & Format$(Precio, "#,##0.00")
         Producto = DatasC.fields("Producto")
         cPrint.letraTipo TipoArialNarrow, 8
         K = Len(Producto)
         Do While (cPrint.anchoTexto(MidStrg(Producto, 1, K)) >= EtiquetaAncho) And (K > 0)
            K = K - 1
         Loop
         K = K - 1
         Producto = MidStrg(Producto, 1, K)
        'Imprimir Etiqueta
         cPrint.generarBarras CodigoBarra, cC128_A, PosIzq + 0.2, PosSup + 0.2, 10, 1
         cPrint.printTexto PosIzq + 0.2, PosSup + 1.5, Producto
         cPrint.printTexto PosIzq + 0.2, PosSup + 1.8, "*" & NombreComercial & "*"
         PosIzq = PosIzq + EtiquetaAncho
         I = I + 1
         If I > CantEtiquetaX Then
            I = 1
            PosIzq = 0.6
            PosSup = PosSup + EtiquetaAlto
            J = J + 1
         End If
         If J > CantEtiquetaY Then
            cPrint.paginaNueva
            J = 1
         End If
         Progreso_Barra.Mensaje_Box = CodigoBarra & " => " & Producto
         Progreso_Esperar
         DatasC.MoveNext
      Loop
      RatonNormal
      MensajeEncabData = ""
      cPrint.finalizaImpresion
      Progreso_Final
   End If
Exit Sub
Errorhandler:
    RatonNormal
    cPrint.finalizaImpresion
    Progreso_Final
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Codigos_De_Barras(Cantidad As Integer, _
                                      CodigoBarra As String)
Dim DatasC As ADODB.Recordset
Dim SizeLetra As Integer
Dim PosSup, PosIzq

Dim PDF_Titulo As String
Dim PDF_Nombre_Documento As String

SizeLetra = 6
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   InicioX = 1: InicioY = 0
   Progreso_Iniciar
  'Geneeramos el documento
   
   tPrint.TipoImpresion = Es_Printer
   tPrint.NombreArchivo = "Etiquetad codigos de barra"
   tPrint.TituloArchivo = "Estanteria de Codigos de Barra"
   tPrint.OrientacionPagina = Orientacion_Pagina
   tPrint.TipoLetra = TipoTimes
   tPrint.PorteLetra = 8
   tPrint.VerDocumento = True
   tPrint.PaginaA4 = True
   tPrint.EsCampoCorto = True
   
   Set cPrint = New cImpresion
   cPrint.iniciaImpresion

   sSQL = "SELECT * " _
        & "FROM Catalogo_Productos " _
        & "WHERE Codigo_Barra = '" & CodigoBarra & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' "
   Select_AdoDB DatasC, sSQL
   If DatasC.RecordCount > 0 Then
   
      Precio = DatasC.fields("PVP")
      If DatasC.fields("IVA") Then Precio = Precio + (Precio * Porc_IVA)
      'MsgBox Precio
      CodigoP = " $ " & Format$(Precio, "#,##0.00")
      Producto = DatasC.fields("Producto")
      cPrint.tipoNegrilla = True
      cPrint.colorDeLetra = Negro
      If Cantidad >= 1 Then
         Pagina = 1: PosLinea = 1: PosColumna = 1
         cPrint.letraTipo TipoArialNarrow, 8
         J = Len(Producto)
         Do While (cPrint.anchoTexto(MidStrg(Producto, 1, J)) > 4.5) And (J > 0)
            J = J - 1
         Loop
         Producto = MidStrg(Producto, 1, J)
         cPrint.letraTipo TipoArialNarrow, 9
         J = 30
         Do While (cPrint.anchoTexto(MidStrg(CodigoBarra & Space(J) & CodigoP, 1, J)) > 2.5) And (J > 0)
            J = J - 1
         Loop
         'CodigoBarra = CodigoBarra & Space(J) & CodigoP
         Cadena = CodigoBarra & Space(J) & CodigoP
         Contador = 1
         For I = 1 To Cantidad
            If Contador > 3 Then
               Contador = 1
               PosColumna = 1
               PosLinea = PosLinea + 2.1
            End If
            If PosLinea >= 26 Then
               cPrint.paginaNueva
               Pagina = 1
               PosLinea = 1
               Contador = 1
               PosColumna = 1
            End If
            cPrint.generarBarras CodigoBarra, cC128_B, PosColumna, PosLinea, 13, 1, False
            PosIzq = PosColumna
            PosSup = PosLinea + 1.1
            If PosIzq <= 0 Then PosIzq = PosColumna
            If PosSup <= 0 Then PosSup = PosLinea
            
            cPrint.letraTipo TipoArial, 5
            cPrint.printTexto PosIzq, PosSup, Empresa & "     *" & NombreComercial & "*"
            PosSup = PosSup + cPrint.altoTexto("H")
            
            cPrint.letraTipo TipoArial, 8
            cPrint.printTexto PosIzq, PosSup, Cadena
             
            PosSup = PosSup + cPrint.altoTexto("H")
            cPrint.printTexto PosIzq, PosSup, Producto
            Contador = Contador + 1
            PosColumna = PosColumna + 6.7
            Progreso_Barra.Mensaje_Box = CStr(I) & " de " & CStr(Cantidad) & ": " & CodigoBarra
            Progreso_Esperar
         Next I
         RatonNormal
         MensajeEncabData = ""
      End If
      cPrint.finalizaImpresion
      Progreso_Final
   End If
Exit Sub
Errorhandler:
    RatonNormal
    cPrint.finalizaImpresion
    Progreso_Final
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Texo_Impresora(TextoImprimir As String)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   If Len(TextoImprimir) > 1 Then
      InicioX = 1: InicioY = 1
      Progreso_Iniciar
      
     'Geneeramos el documento
      tPrint.TipoImpresion = Es_Printer
      tPrint.NombreArchivo = "Archivo_Texto"
      tPrint.TituloArchivo = "Archivo Texto"
      tPrint.OrientacionPagina = Orientacion_Pagina
      tPrint.TipoLetra = TipoCourierNew
      tPrint.PorteLetra = 8
      tPrint.VerDocumento = True
      tPrint.PaginaA4 = True
      tPrint.EsCampoCorto = True
      Set cPrint = New cImpresion
      cPrint.iniciaImpresion
     
      MensajeEncabData = "IMPRESION DE UN TEXTO"
     
     'Iniciamos la impresion
      cPrint.colorDeLetra = Negro
      cPrint.printTexto 0.01, 1, TextoImprimir
        
      MensajeEncabData = ""
      cPrint.finalizaImpresion
   End If
   Exit Sub
Errorhandler:
    RatonNormal
    cPrint.finalizaImpresion
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Codigos_Barras_Kardex(PictBar As PictureBox, _
                                          CantidadEtiquetas As Integer)
Dim SizeLetra As Integer
Dim PosSup, PosIzq
Dim CodigoBarra As String
Dim CantEtiquetaX As Byte
Dim CantEtiquetaY As Byte
Dim EtiquetaAncho As Single
Dim EtiquetaAlto As Single
Dim CantCodigos As Integer
SizeLetra = 6
CantEtiquetaX = 5
CantEtiquetaY = 10
EtiquetaAncho = 3.8
EtiquetaAlto = 2.5

On Error GoTo Errorhandler
RatonReloj
CantCodigos = InputBox("CANTIDAD DE CODIGOS A IMPRIMIR:", "IMPRESION DE CODIGOS DE BARRAS", CStr(CantidadEtiquetas))
 
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   InicioX = 1: InicioY = 0
   Escala_Centimetro 1, TipoTimes, 10
   If CantCodigos > 0 Then
      Printer.FontBold = True
      Contador = 1
      Pagina = 1
      PosLinea = 1
      Contador = 1
      PosColumna = 1
      I = 1: J = 1
      PosSup = 1.2
      PosIzq = 0.6
      Do While Contador <= CantCodigos
         CodigoBarra = DatInv.Codigo_Barra
         Precio = DatInv.PVP
         If DatInv.IVA Then Precio = Precio + (Precio * Porc_IVA)
         CodigoP = "PVP $ " & Format$(Precio, "#,##0.00")
         Producto = DatInv.Producto
         Printer.FontName = TipoArialNarrow
         Printer.ForeColor = 0
         Printer.FontSize = 8
         K = Len(Producto)
         Do While (Printer.TextWidth(MidStrg(Producto, 1, K)) >= EtiquetaAncho) And (K > 0)
            K = K - 1
         Loop
         K = K - 1
         Producto = MidStrg(Producto, 1, K)
'''         Code39.AlturaBarra = 25
'''         Code39.TamBarra = 1
'''         Code39.ColorCodigo = "N"
'''         Code39.ValorCodigo = CodigoBarra
'''         Code39.RealizarCodigo
'''        'Imprimir Etiqueta
'''         PictBar.Cls
'''         PictBar.AutoRedraw = True
'''         PictBar.FontBold = False
'''         PictBar.Picture = Clipboard.GetData
'''         Printer.DrawWidth = 1
         'Printer.Line (PosIzq + 0.1, PosSup + 0.1)-(PosIzq + EtiquetaAncho - 0.1, PosSup + EtiquetaAlto - 0.1), Negro, B
         Printer.PaintPicture PictBar.Image, PosIzq + 0.2, PosSup + 0.2, EtiquetaAncho - 0.4
        
         Printer.FontSize = 8
         Printer.CurrentX = PosIzq + 0.2
         Printer.CurrentY = PosSup + 1
         Printer.Print CodigoBarra & " (" & Format(Contador, "0000") & ")"
         
         Printer.CurrentX = PosIzq + 0.2
         Printer.CurrentY = PosSup + 1.4
         Printer.Print Producto
         Printer.CurrentX = PosIzq + 0.2
         Printer.CurrentY = PosSup + 1.8
         Printer.Print "*" & NombreComercial & "*"
         PosIzq = PosIzq + EtiquetaAncho
         I = I + 1
         If I > CantEtiquetaX Then
            I = 1
            PosIzq = 0.6
            PosSup = PosSup + EtiquetaAlto
            J = J + 1
         End If
         If J > CantEtiquetaY Then
            Printer.NewPage
            I = 1
            J = 1
            PosSup = 1.2
            PosIzq = 0.6
         End If
         Contador = Contador + 1
      Loop
      RatonNormal
      MensajeEncabData = ""
      Printer.EndDoc
   End If

Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub MuestraError(TituloError, NErr, TError)
  If NErr <> 0 Then MsgBox NErr & vbCrLf & TError, vbInformation, TituloError
End Sub

'''Public Function PapelesDeImpresora(NombreImpresora As String, LstPRN() As String) As Long
'''Dim I As Long
'''Dim Ret As Long
'''Dim PaperNo() As Integer
'''Dim PaperName() As String
'''Dim PaperSize() As POINTS
'''Dim CadAuxPRN As String
'''RatonReloj
'''PapelesDeImpresora = 0
''''Total papeles que admite la impresora
'''Ret = DeviceCapabilities(NombreImpresora, "LPT1", DC_PAPERS, ByVal 0&, ByVal 0&)
'''If Ret < 1 Then GoTo ErrorPapelesImpresora
'''ReDim PaperNo(1 To Ret) As Integer
'''Call DeviceCapabilities(NombreImpresora, "LPT1", DC_PAPERS, PaperNo(1), ByVal 0&)
'''
''''Nº, Nombre, Alto y ancho de los papeles
'''Dim arrPageName() As Byte
'''Dim allNames As String
'''Dim lStart As Long, lEnd As Long
'''ReDim PaperName(1 To Ret) As String
'''ReDim arrPageName(1 To Ret * 64) As Byte
'''
'''Call DeviceCapabilities(NombreImpresora, "LPT1", DC_PAPERNAMES, arrPageName(1), ByVal 0&)
'''allNames = StrConv(arrPageName, vbUnicode)
'''
'''I = 1
'''Do
'''    lEnd = InStr(lStart + 1, allNames, Chr$(0), vbBinaryCompare)
'''    If (lEnd > 0) And (lEnd - lStart - 1 > 0) Then
'''         PaperName(I) = MidStrg(allNames, lStart + 1, lEnd - lStart - 1)
'''        I = I + 1
'''    End If
'''    lStart = lEnd
'''Loop Until lEnd = 0
'''
''''Tamaño papeles
'''ReDim PaperSize(1 To Ret) As POINTS
'''
'''Call DeviceCapabilities(NombreImpresora, "LPT1", DC_PAPERSIZE, PaperSize(1), ByVal 0&)
''''Introducimos en el List
'''ReDim LstPRN(Ret) As String
'''   For I = 1 To Ret
'''
'''       MsgBox "=======>>>   " & NombreImpresora & vbCrLf & PaperNo(I)
'''
'''       LstPRN(I) = PaperNo(I) & " - " & PaperName(I) _
'''                 & " (" & Format$(PaperSize(I).X / 100, "#,##0.0") & " x " _
'''                 & Format$(PaperSize(I).Y / 100, "#,##0.0") & ")"
'''       If I = 1 Then LstPRN(0) = LstPRN(1)
'''       If PaperNo(I) = 9 Then
'''          LstPRN(0) = PaperNo(I) & " - " & PaperName(I) _
'''                    & " (" & Format$(PaperSize(I).X / 100, "#,##0.0") & " x " _
'''                    & Format$(PaperSize(I).Y / 100, "#,##0.0") & ")"
'''       End If
'''   Next I
''''LstPRN.Text = CadAuxPRN
'''PapelesDeImpresora = Ret
'''RatonNormal
'''Exit Function
'''ErrorPapelesImpresora:
'''     RatonNormal
'''     PapelesDeImpresora = 0
'''     MuestraError "PapelesImpresora", Err, Error
'''End Function

'''Sub Leer_Lista_De_Impresoras()
''' 'Revizamos cuantas Impresoras estan instaladas y activas
'''  ReDim ListaDeImpresoras(Printers.Count + 1) As String
'''  For I = 0 To Printers.Count - 1
'''      ListaDeImpresoras(I) = Printers(I).DeviceName
'''  Next I
'''  ListaDeImpresoras(Printers.Count) = Impresota_PDF
'''End Sub

Public Function PapelesImpresora(NombreImpresora As String, LstPRN As ListBox) As Boolean
Dim I As Long
Dim ret As Long
Dim PaperNo() As Integer
Dim PaperName() As String
Dim PaperSize() As POINTS
Dim CadAuxPRN As String
'Total papeles que admite la impresora
ret = DeviceCapabilities(NombreImpresora, "LPT1", DC_PAPERS, ByVal 0&, ByVal 0&)
If ret < 1 Then GoTo ErrorPapelesImpresora
ReDim PaperNo(1 To ret) As Integer
Call DeviceCapabilities(NombreImpresora, "LPT1", DC_PAPERS, PaperNo(1), ByVal 0&)

'Nº, Nombre, Alto y ancho de los papeles
Dim arrPageName() As Byte
Dim allNames As String
Dim lStart As Long, lEnd As Long
ReDim PaperName(1 To ret) As String
ReDim arrPageName(1 To ret * 64) As Byte

Call DeviceCapabilities(NombreImpresora, "LPT1", DC_PAPERNAMES, arrPageName(1), ByVal 0&)
allNames = StrConv(arrPageName, vbUnicode)

I = 1
Do
    lEnd = InStr(lStart + 1, allNames, Chr$(0), vbBinaryCompare)
    If (lEnd > 0) And (lEnd - lStart - 1 > 0) Then
         PaperName(I) = MidStrg(allNames, lStart + 1, lEnd - lStart - 1)
        I = I + 1
    End If
    lStart = lEnd
Loop Until lEnd = 0

'Tamaño papeles
ReDim PaperSize(1 To ret) As POINTS
Call DeviceCapabilities(NombreImpresora, "LPT1", DC_PAPERSIZE, PaperSize(1), ByVal 0&)
'Introducimos en el List
LstPRN.Clear
If NombreImpresora = Impresota_PDF Then
   LstPRN.AddItem "009 - A4 (21.0 x 29.7)"
Else
   For I = 1 To ret
       LstPRN.AddItem Format$(PaperNo(I), "000") & " - " & PaperName(I) _
                      & " (" & Format$(PaperSize(I).X / 100, "#,##0.0") & " x " _
                      & Format$(PaperSize(I).y / 100, "#,##0.0") & ")"
   Next I
End If
'MsgBox "PRN: " & LstPRN.Text & " ===> " & CadAuxPRN
Exit Function
ErrorPapelesImpresora:
 MuestraError "PapelesImpresora", Err, Error
End Function

Public Function PonImpresoraDefecto(NombreImpresora As String) As Boolean
'Función que establece una impresora determinada
Dim I As Integer
On Error GoTo ErrorPonImpresoraDefecto
PonImpresoraDefecto = False
If NombreImpresora <> "" Then
  'Recorro la coleccion de impresoras activas en cad puesto
   For I = 0 To Printers.Count - 1
     'Miro si la impresora que busco esta en la coleccion
      If Printers(I).DeviceName = NombreImpresora Then
          'Si es la que busco la establezco
          Set Printer = Printers(I)
          Exit For
      End If
   Next I
   PonImpresoraDefecto = True
End If
Exit Function
ErrorPonImpresoraDefecto:
     MuestraError "PonImpresoraDefecto", Err, Error
End Function

Public Sub AbrirArchivoTexto(TextFile As String)
Dim NumFile As Long
Dim LineaDeTexto As String
  RatonReloj
  TextoFile = " " & vbCrLf
  If Len(TextFile) > 1 Then
     NumFile = FreeFile
     Open TextFile For Input As #NumFile
     Do While Not EOF(NumFile)
        Line Input #NumFile, LineaDeTexto
        TextoFile = TextoFile & LineaDeTexto & " " & vbCrLf
     Loop
     Close #NumFile
  End If
  TextoFile = TextoFile & " " & vbCrLf
  'MsgBox TextoFile
  RatonNormal
End Sub

Public Sub AbrirArchivo(TextFile As TextBox)
Dim NumFile As Long
  RatonReloj
  If Len(RutaGeneraFile) > 1 Then
  TextFile.Text = ""
  NumFile = FreeFile
  Open RutaGeneraFile For Input As #NumFile
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cadena
     TextFile.Text = TextFile.Text & Cadena & vbCrLf
  Loop
  Close #NumFile
  End If
  RatonNormal
End Sub

Public Sub Abrir_Excel(FileExcel As String)
Dim iRet As Long
  If Dir$(FileExcel) <> "" Then
     iRet = Shell("rundll32.exe url.dll,FileProtocolHandler " & FileExcel, vbMaximizedFocus)
  Else
     MsgBox "EL ARCHIVO:" & vbCrLf & vbCrLf & FileExcel & vbCrLf & vbCrLf & "NO EXISTE O ESTA DAÑADO"
  End If
End Sub

Public Sub Presenta_Archivo_PDF(PathFile As String)
Dim Results As String
  If Len(PathFile) > 1 Then
     Results = Dir$(PathFile)
     If Results <> "" Then
        RatonReloj
        Results = Shell("rundll32.exe url.dll,FileProtocolHandler " & PathFile, vbMaximizedFocus)
        RatonNormal
     End If
  End If
End Sub

Public Sub TextoValidoVar(TextB As String, _
                          Optional Numerio As Boolean, _
                          Optional Mayusculas As Boolean)
    If IsNull(TextB) Then TextB = ""
    TextB = Replace(TextB, vbCr, "")
    TextB = Replace(TextB, vbLf, "")
    If TextB = "" Then
       If Numerio Then TextB = "0" Else TextB = Ninguno
    End If
    If Mayusculas Then TextB = UCaseStrg(TextB)
    TextB = TrimStrg(TextB)
End Sub

''Public Sub Texto_UCaseStrg(TextB As TextBox)
''   TextB.Text = UCaseStrg(TrimStrg(TextB.Text))
''End Sub

Public Sub LimpiarTexto(TextB As TextBox, _
                        Limpiar As Boolean)
  If Limpiar Then TextB.Text = "" Else MarcarTexto TextB
End Sub

Public Function TextoMaximo(TextB As TextBox) As Boolean
Dim TextoMaximo1 As Boolean
    TextoMaximo1 = False
    If TextB.MaxLength <> 0 And Len(TextB.Text) >= TextB.MaxLength Then
       TextoMaximo1 = True
    End If
    TextoMaximo = TextoMaximo1
End Function

Public Sub SiguienteControl()
   Pulsar_Tecla (vbKeyTab)
   'SendKeys "{TAB}", False
End Sub

Public Sub LimpiarLinea(x1 As Single, _
                        y1 As Single, _
                        Optional PonerLineas As Boolean, _
                        Optional Color_Fondo_Texto As Long)
Dim ImpBarraB As Boolean
  If Color_Fondo_Texto = 0 Then Color_Fondo_Texto = Blanco
  If y1 < LimiteAlto Then
  If x1 > 0 And y1 > 0 Then
     ImpBarraB = False
     If AltoLetra < 0.35 Then AltoLetra = 0.35
     If LongNumero > 0 Then
        If (x1 - 0.05) > (x1 + LongNumero) Then ImpBarraB = True
        If ImpBarraB Then Printer.Line (x1 - 0.05, y1)-(x1 + LongNumero + 0.1, y1 + AltoLetra), Color_Fondo_Texto, BF
     Else
        If (x1 - 0.05) > (x1 + AnchoPapel) Then ImpBarraB = True
        If ImpBarraB Then Printer.Line (x1 - 0.05, y1)-(x1 + AnchoPapel, y1 + AltoLetra), Color_Fondo_Texto, BF
     End If
  End If
  If PonerLineas Then Printer.Line (x1, y1 - 0.05)-(x1, y1 + AltoLetra + 0.05), Negro
  'MsgBox "Linea: " & LongNumero
  LongNumero = 0
  End If
End Sub

Public Sub LimpiarLineaTexto(x1 As Single, _
                             y1 As Single, _
                             LineaTexto As String, _
                             Optional Color_Fondo_Texto As Long)
Dim anchoTexto As Single
If Color_Fondo_Texto = 0 Then Color_Fondo_Texto = Blanco
  If x1 > 0 And y1 > 0 Then
     anchoTexto = Printer.TextWidth(LineaTexto)
     AltoLetra = Printer.TextHeight("H")
     Printer.Line (x1 + 0.1, y1)-(x1 + anchoTexto + 0.1, y1 + AltoLetra), Color_Fondo_Texto, BF
  End If
End Sub

Public Sub FormatoObjeto(Objeto As Control, _
                         Tipo)
  Select Case Tipo
    Case vbByte, vbInteger, vbLong
         If Objeto.Text = "" Then Objeto.Text = "0"
         Objeto.Text = Format$(Val(Objeto.Text), "##0.00")
    Case vbDouble
         If Objeto.Text = "" Then Objeto.Text = "0"
         Objeto.Text = Format$(Val(Objeto.Text), "##0.00%")
    Case vbSingle, vbCurrency
         If Objeto.Text = "" Then Objeto.Text = "0"
         Objeto.Text = Format$(Val(Objeto.Text), "#,##0.00")
    Case Else
         If Objeto.Text = "" Then Objeto.Text = Ninguno
  End Select
End Sub

Public Function FechaStrgDias(Fechas As String) As String
Dim dd, MM, AA As String
Dim DiasSem As String
  If Fechas = "00/00/0000" Then Fechas = FechaSistema
  If IsNull(Fechas) Or IsEmpty(Fechas) Then
     FechaStrgDias = "Ninguno"
  Else
  DiasSem = ""
  Fechas = Format$(Fechas, FormatoFechas)
  Select Case Weekday(Fechas)
    Case 1: DiasSem = "Domingo"
    Case 2: DiasSem = "Lunes"
    Case 3: DiasSem = "Martes"
    Case 4: DiasSem = "Miércoles"
    Case 5: DiasSem = "Jueves"
    Case 6: DiasSem = "Viernes"
    Case 7: DiasSem = "Sábado"
  End Select
  dd = Format$(Day(Fechas), "00")
  MM = Format$(Month(Fechas), "00")
  AA = Format$(Year(Fechas), "0000")
  If (Val(AA) < 1) Then AA = Anio
  If ((Val(MM) < 1) Or (Val(MM) > 12)) Then MM = Mes
  If (Val(dd) < 1) Or (Val(dd) > MaximoDia(MM, AA)) Then dd = MaximoDia(MM, AA)
  FechaStrgDias = DiasSem & ", " & dd & "/" & MidStrg(UCaseStrg(MesesLetras(Val(MM))), 1, 3) & "/" & AA
  End If
End Function

Public Function FechaDiaSem(Fechas As String) As String
  Fechas = Format$(Fechas, FormatoFechas)
  Select Case Weekday(Fechas)
    Case 1: FechaDiaSem = "Domingo"
    Case 2: FechaDiaSem = "Lunes"
    Case 3: FechaDiaSem = "Martes"
    Case 4: FechaDiaSem = "Miércoles"
    Case 5: FechaDiaSem = "Jueves"
    Case 6: FechaDiaSem = "Viernes"
    Case 7: FechaDiaSem = "Sábado"
  End Select
End Function

Public Sub FechaValida(NomBox As MaskEdBox, Optional ChequearCierreMes As Boolean)
Dim AdoCierre As ADODB.Recordset
Dim ErrorFecha As Boolean
Dim sSQL1 As String
Dim FechaFin1 As String
Dim TipoError As String
Dim AnioFecha As String

 'Empezamos a verificar la fecha ingresada
  RatonReloj
  TipoError = ""
  ErrorFecha = False
  If NomBox.Text = LimpiarFechas Then NomBox.Text = FechaSistema
  NomBox.Text = Format$(NomBox.Text, FormatoFechas)
  If IsDate(NomBox.Text) Then
    'Averiguamos si esta cerrado el mes de procesamiento
     FechaCierre = "01/" & Month(FechaSistema) & "/" & Year(FechaSistema)
     AnioFecha = Format(Year(NomBox.Text), "0000")
     If Val(AnioFecha) >= 1900 Then
        FechaFin1 = BuscarFecha(NomBox.Text)
        sSQL1 = "SELECT Fecha_Inicial " _
              & "FROM Fechas_Balance " _
              & "WHERE Periodo = '" & Periodo_Contable & "' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Cerrado = " & Val(adFalse) & " " _
              & "AND #" & FechaFin1 & "# BETWEEN Fecha_Inicial AND Fecha_Final " _
              & "AND Detalle LIKE '" & AnioFecha & "%' "
        Select_AdoDB AdoCierre, sSQL1
        If AdoCierre.RecordCount > 0 Then FechaCierre = AdoCierre.fields("Fecha_Inicial")
        AdoCierre.Close
       'MsgBox ChequearCierreMes & vbCrLf & ErrorFecha
        If ChequearCierreMes Then
           If CFechaLong(NomBox.Text) >= CFechaLong(FechaCierreFiscal) Then
              If CFechaLong(NomBox.Text) < CFechaLong(FechaCierre) Then
                 ErrorFecha = True
                 TipoError = TipoError & "ESTA BLOQUEADA POR EL CIERRE DEL MES, SOLICITE A CONTABILIDAD QUE LE ACTIVE"
              End If
           Else
              ErrorFecha = True
              TipoError = TipoError & "ES INFERIOR AL CIERRE FISCAL: " & FechaCierreFiscal
           End If
        End If
     Else
        ErrorFecha = True
        TipoError = TipoError & "ES MENOR A 1900"
     End If
  Else
     ErrorFecha = True
     TipoError = TipoError & "ES INCORRECTA, VUELVA A INGRESAR"
  End If
 'Resultado Final de la verificacion de la Fecha ingresada
  RatonNormal
  If ErrorFecha Then
     MsgBox "LA FECHA QUE ESTA INTENTANDO INGRESAR " & TipoError
     NomBox.Text = LimpiarFechas
     NomBox.SetFocus
  End If
End Sub

Public Sub Validar_Porc_IVA(FechaIVA As String)
Dim AdoCierre As ADODB.Recordset
Dim sSQL1 As String
    RatonReloj
   'Carga la Tabla de Porcentaje IVA
    If FechaIVA = "00/00/0000" Then FechaIVA = FechaSistema
    sSQL1 = "SELECT * " _
          & "FROM Tabla_Por_ICE_IVA " _
          & "WHERE IVA <> " & Val(adFalse) & " " _
          & "AND Fecha_Inicio <= #" & BuscarFecha(FechaIVA) & "# " _
          & "AND Fecha_Final >= #" & BuscarFecha(FechaIVA) & "# " _
          & "ORDER BY Porc DESC "
    Select_AdoDB AdoCierre, sSQL1
    If AdoCierre.RecordCount > 0 Then
       Porc_IVA = Redondear(AdoCierre.fields("Porc") / 100, 2)
       Cod_Porc_IVA = AdoCierre.fields("Codigo")
    End If
    AdoCierre.Close
   'MsgBox "===--->> " & Porc_IVA
    RatonNormal
End Sub

Public Sub Obtener_Porc_IVA(FechaIVA As String, CodPorcIva As Byte)
Dim AdoCierre As ADODB.Recordset
Dim sSQL1 As String
    RatonReloj
   'Carga la Tabla de Porcentaje IVA
    If FechaIVA = "00/00/0000" Then FechaIVA = FechaSistema
    sSQL1 = "SELECT * " _
          & "FROM Tabla_Por_ICE_IVA " _
          & "WHERE IVA <> " & Val(adFalse) & " " _
          & "AND Fecha_Inicio <= #" & BuscarFecha(FechaIVA) & "# " _
          & "AND Fecha_Final >= #" & BuscarFecha(FechaIVA) & "# " _
          & "AND Codigo = " & CodPorcIva & " " _
          & "ORDER BY Porc DESC "
    Select_AdoDB AdoCierre, sSQL1
    If AdoCierre.RecordCount > 0 Then Porc_IVA = Redondear(AdoCierre.fields("Porc") / 100, 2)
    AdoCierre.Close
    RatonNormal
End Sub

Public Sub Obtener_Cod_Porc_IVA(FechaIVA As String, PorcIva As Byte)
Dim AdoCierre As ADODB.Recordset
Dim sSQL1 As String
    RatonReloj
   'Carga la Tabla de Porcentaje IVA
    If FechaIVA = "00/00/0000" Then FechaIVA = FechaSistema
    sSQL1 = "SELECT * " _
          & "FROM Tabla_Por_ICE_IVA " _
          & "WHERE IVA <> " & Val(adFalse) & " " _
          & "AND Fecha_Inicio <= #" & BuscarFecha(FechaIVA) & "# " _
          & "AND Fecha_Final >= #" & BuscarFecha(FechaIVA) & "# " _
          & "AND Porc = " & PorcIva & " " _
          & "ORDER BY Porc DESC "
    Select_AdoDB AdoCierre, sSQL1
    If AdoCierre.RecordCount > 0 Then
       Cod_Porc_IVA = AdoCierre.fields("Codigo")
       Porc_IVA = Redondear(AdoCierre.fields("Porc") / 100, 2)
    End If
    AdoCierre.Close
    RatonNormal
End Sub

Public Sub Crear_Clientes(Codigo_Acceso As String)
Dim SQLAux As String
Dim AdoDBAcceso As ADODB.Recordset
   SQLAux = "SELECT * " _
          & "FROM Clientes " _
          & "WHERE Codigo = '" & Codigo_Acceso & "' "
   Select_AdoDB AdoDBAcceso, SQLAux
  'MsgBox AdoDBAcceso.RecordCount
   If AdoDBAcceso.RecordCount <= 0 Then
      SetAdoAddNew "Accesos", True
      SetAdoFields "Codigo", Codigo_Acceso
      Select Case Codigo_Acceso
        Case "ACCESO01"
             SetAdoFields "Usuario", "Supervisor"
             SetAdoFields "Nombre_Completo", "Supervidor General"
        Case "ACCESO02"
             SetAdoFields "Usuario", "Administrador"
             SetAdoFields "Nombre_Completo", "Administrador de Red"
        Case "ACCESO03"
             SetAdoFields "Usuario", "Contador"
             SetAdoFields "Nombre_Completo", "Contador General"
        Case "ACCESO04"
             SetAdoFields "Usuario", "Auxiliar"
             SetAdoFields "Nombre_Completo", "Auxiliar de Agencia"
        Case "ACCESO05"
             SetAdoFields "Usuario", "Gerente"
             SetAdoFields "Nombre_Completo", "Gerente General"
        Case "ACCESO06"
             SetAdoFields "Usuario", "Reabrir"
             SetAdoFields "Nombre_Completo", "Reabrir Periodo"
        Case "ACCESO07"
             SetAdoFields "Usuario", "CierreMes"
             SetAdoFields "Nombre_Completo", "Cierre de Meses"
        Case "0702164179"
             SetAdoFields "Usuario", "Walter"
             SetAdoFields "Nombre_Completo", "Walter Vaca Prieto"
        Case "."
             SetAdoFields "Usuario", "..."
             SetAdoFields "Nombre_Completo", Ninguno
      End Select
      SetAdoFields "Clave", "070216"
      SetAdoFields "TODOS", True
      SetAdoFields "Supervisor", True
      SetAdoUpdate
   End If
   AdoDBAcceso.Close
End Sub

Public Sub Progreso_Iniciar_Errores()
  sSQL = "DELETE * " _
       & "FROM Tabla_Temporal " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Modulo = '" & NumModulo & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
End Sub

Public Sub Progreso_Iniciar(Optional Imprimiendo As Boolean)
Dim ImagenIco As String
    With MDIFormulario
         Procesando = 0
         Progreso_Barra.Incremento = 0
         Progreso_Barra.Puntos = 0
         Progreso_Barra.color = 0
         Color_Fondo .hwnd, RGB(153, 153, 255)  'Magenta
        .StaBarEmp.Panels.Item(5).Bevel = sbrRaised
         If Imprimiendo Then ImagenIco = RutaSistema & "\FORMATOS\Impresora.ico" Else ImagenIco = RutaSistema & "\FORMATOS\Procesando.ico"
         If Dir(ImagenIco) <> "" Then .StaBarEmp.Panels.Item(5).Picture = LoadPicture(ImagenIco) Else .StaBarEmp.Panels.Item(5).Picture = LoadPicture()
         If Progreso_Barra.Valor_Maximo <= 0 Then Progreso_Barra.Valor_Maximo = 100
         If Len(Progreso_Barra.Mensaje_Box) > 1 Then .StaBarEmp.Panels.Item(5).Text = Progreso_Barra.Mensaje_Box Else .StaBarEmp.Panels.Item(5).Text = "Procesando..."
        
        'Actualiza estado de barra de colores
        .ProgressBarEstado.value = Progreso_Barra.Incremento
        .ProgressBarEstado.Max = Progreso_Barra.Valor_Maximo
    End With
End Sub

Public Sub Progreso_Esperar(Optional NoIncrementar As Boolean)
Dim CadProgreso As String
Dim ImagenIco As String
Dim Porcentaje As Long
    DoEvents
    With Progreso_Barra
         RatonReloj
         CadProgreso = ""
         If .Puntos < 0 Then .Puntos = 0
         If .Puntos > 20 Then .Puntos = 0
         If .color > 15 Then .color = 0
         If .Valor_Maximo <= 0 Then .Valor_Maximo = 1
       ' Establece el color del progress
         Color_Progreso MDIFormulario.ProgressBarEstado.hwnd, QBColor(.color)
         'If .Incremento <= 0 Then Progreso_Iniciar
         If .Incremento < .Valor_Maximo Then
             Porcentaje = .Incremento / .Valor_Maximo
             If Porcentaje > 0 Then CadProgreso = Format(Porcentaje, "00%") & ", "
             If Len(.Mensaje_Box) > 1 Then CadProgreso = CadProgreso & .Mensaje_Box Else CadProgreso = CadProgreso & " Espere por favor"
             CadProgreso = CadProgreso & String(.Puntos, ".")
         Else
             If Len(.Mensaje_Box) > 1 Then CadProgreso = .Mensaje_Box
         End If
         If CadProgreso = "" Then CadProgreso = "Espere un momento, procesando..."
         ImagenIco = RutaSistema & "\FORMATOS\Procesando" & Format(Procesando, "00") & ".ico"
         If Dir(ImagenIco) <> "" Then MDIFormulario.StaBarEmp.Panels.Item(5).Picture = LoadPicture(ImagenIco)
         MDIFormulario.StaBarEmp.Panels.Item(5).Text = CadProgreso
         If .Incremento > .Valor_Maximo Then .Incremento = .Valor_Maximo
         If Procesando > 10 Then Procesando = 0
         MDIFormulario.ProgressBarEstado.Max = .Valor_Maximo
         MDIFormulario.ProgressBarEstado.value = .Incremento
         MDIFormulario.StaBarEmp.Refresh
         If Not NoIncrementar Then .Incremento = .Incremento + 1
         Procesando = Procesando + 1
        .Puntos = .Puntos + 1
        .color = .color + 1
     End With
End Sub

Public Sub Progreso_Final()
    'Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
    With MDIFormulario
        .StaBarEmp.Panels.Item(5).Text = ""
        .StaBarEmp.Panels.Item(5).Bevel = sbrInset
        .StaBarEmp.Panels.Item(5).Picture = LoadPicture()
         Color_Fondo .ProgressBarEstado.hwnd, RGB(153, 153, 255)
        .ProgressBarEstado.value = 0 '.ProgressBarEstado.Max
        .StaBarEmp.Panels.Item(5).Picture = LoadPicture(RutaSistema & "\FORMATOS\Visto.ico")
    End With
    RatonNormal
End Sub

Public Sub PonerDirEmpresa()
Dim AnchoPantalla As Single
Dim AnchoTemp As Single
Dim Result As Long
  ContadorEstados = 0
  If Val(NumEmpresa) > 0 Then
     RatonReloj
     
'''     Ambiente = Leer_Campo_Empresa("Ambiente")
'''     Obligado_Conta = Leer_Campo_Empresa("Obligado_Conta")
'''     ContEspec = Leer_Campo_Empresa("Codigo_Contribuyente_Especial")
    
    'Informar Modo de Ambiente para comprobantes electronicos
     Select Case Ambiente
       Case "1": MDIFormulario.MAmbiente.Caption = "AMBIENTE DE PRUEBA"
       Case "2": MDIFormulario.MAmbiente.Caption = "AMBIENTE EN PRODUCCION"
       Case Else: MDIFormulario.MAmbiente.Caption = ""
     End Select

    'Informamos los datos de la empresa
     RatonReloj
     Progreso_Barra.Incremento = 0
     Progreso_Barra.Valor_Maximo = 100
     Progreso_Barra.Mensaje_Box = "Conectandose a la base de Datos"
     
     MDIFormulario.Caption = Empresa & ":"
     MDIFormulario.StaBarEmp.Font.bold = True
     MDIFormulario.StaBarEmp.Font.Size = 9
     MDIFormulario.StaBarEmp.Font.Name = TipoArialNarrow
     
     Codigo = NumEmpresa & ".- " & Modulo
     MDIFormulario.StaBarEmp.Panels.Item(1).MinWidth = 3300 'MDIFormulario.TextWidth(Codigo) * 2 '2300
     MDIFormulario.StaBarEmp.Panels.Item(1).Text = Codigo
     Codigo = ULCase(NombreUsuario)
     MDIFormulario.StaBarEmp.Panels.Item(2).MinWidth = 3200 'MDIFormulario.PictMDI.TextWidth(Codigo) * 2 '2900
     MDIFormulario.StaBarEmp.Panels.Item(2).Text = Codigo
     If Periodo_Contable <> Ninguno Then Codigo = "PERIODO: " & Periodo_Contable Else Codigo = "PERIODO ACTUAL"
     MDIFormulario.StaBarEmp.Panels.Item(3).MinWidth = 2400 'MDIFormulario.PictMDI.TextWidth(Codigo) * 2 '2100
     MDIFormulario.StaBarEmp.Panels.Item(3).Text = Codigo
     If SQL_Server Then Codigo = "SQL Server" Else Codigo = "My SQL"
     MDIFormulario.StaBarEmp.Panels.Item(4).MinWidth = 1700 'MDIFormulario.PictMDI.TextWidth(Codigo) * 2 '1500
     MDIFormulario.StaBarEmp.Panels.Item(4).Text = Codigo
     AnchoTemp = 0
     Cadena = ""
     For I = 1 To 4
         Cadena = Cadena & MDIFormulario.StaBarEmp.Panels.Item(I).Text & vbTab & vbTab & vbTab & " = " & MDIFormulario.StaBarEmp.Panels.Item(I).MinWidth & vbCrLf
         AnchoTemp = AnchoTemp + MDIFormulario.StaBarEmp.Panels.Item(I).MinWidth
     Next I
    'MsgBox Cadena
     AnchoPantalla = Screen.width - 50 '
    'MDIFormulario.StaBarEmp.Panels.Item(6).AutoSize = sbrSpring
     If AnchoPantalla > AnchoTemp Then MDIFormulario.StaBarEmp.Panels.Item(5).MinWidth = Redondear(AnchoPantalla - AnchoTemp)
     MDIFormulario.StaBarEmp.Panels.Item(5).Text = ""
     MDIFormulario.StaBarEmp.Panels.Item(5).Picture = LoadPicture("")
     
     Color_Fondo MDIFormulario.ProgressBarEstado.hwnd, RGB(153, 153, 255)
     EmpresaActual = "[" & RutaEmpresa & "]."
    
    'Leer_Lista_De_Impresoras
               
    'Seteamos datos iniciales de las cuentas y codigos del sistema
     Progreso_Barra.Mensaje_Box = "Verificando Datos iniciales de la Empresa"
     
     RatonReloj
     
    'Verificamos si existe la carpeta en red de los comprobantes electronicos
     If Modulo <> "UPDATE" Then
        RatonReloj
        Progreso_Barra.Mensaje_Box = "Extrayendo Certificado"
        
       'Creamos carpetas si no existen
'''        If Not Existe_Carpeta(WindowsDirectory & "\LIBRERIA") Then MkDir (WindowsDirectory & "\LIBRERIA")
        If Not Existe_Carpeta(RutaSysBases & "\TEMP") Then MkDir RutaSysBases & "\TEMP"
        
       'Verificamos y creamos carpetas de firma electronica
        RutaDocumentos = RutaSysBases & "\CE"
        If Not Existe_Carpeta(RutaDocumentos) Then MkDir RutaDocumentos
       
        RutaDocumentos = RutaSysBases & "\CE\CE" & NumEmpresa
        If Not Existe_Carpeta(RutaDocumentos) Then MkDir RutaDocumentos
        If Not Existe_Carpeta(RutaDocumentos & "\Comprobantes Autorizados") Then MkDir (RutaDocumentos & "\Comprobantes Autorizados")
        If Not Existe_Carpeta(RutaDocumentos & "\Comprobantes Firmados") Then MkDir (RutaDocumentos & "\Comprobantes Firmados")
        If Not Existe_Carpeta(RutaDocumentos & "\Comprobantes Generados") Then MkDir (RutaDocumentos & "\Comprobantes Generados")
        If Not Existe_Carpeta(RutaDocumentos & "\Comprobantes no Autorizados") Then MkDir (RutaDocumentos & "\Comprobantes no Autorizados")
       
       'Instalando el certificado de la firma electronica
'''        If RutaCertificado <> Ninguno Then
'''           If Not Existe_File(WindowsDirectory & "\LIBRERIA\" & RutaCertificado) Then
'''              Result = Copiar_Archivos(RutaSistema & "\CERTIFIC\", WindowsDirectory & "\LIBRERIA\", RutaCertificado)
'''           End If
'''        End If
       'Averiguamos si las librerias estan instaladas
'''        If Not Existe_File(WindowsDirectory & "\LIBRERIA\FacturacionElect.tlb") Then
'''           Result = Copiar_Archivos(RutaSistema & "\LIBRERIA\", WindowsDirectory & "\LIBRERIA\", "*.dll")
'''           Result = Copiar_Archivos(RutaSistema & "\LIBRERIA\", WindowsDirectory & "\LIBRERIA\", "*.bat")
'''           MsgBox "COMUNIQUESE AL PROVEEDOR DEL PROGRAMA," & vbCrLf _
'''                & "SU SISTEMA NECESITA UNA ACTUALIZACION," & vbCrLf _
'''                & "DE " & WindowsDirectory & "\LIBRERIA\FacturacionElect.tlb" _
'''                & "LOS CONTACTOS TELEFONICOS SON: " & vbCrLf _
'''                & "09-8910-5300/09-9965-4196"
'''        End If
     End If
     RatonNormal
    'MsgBox Empresa & vbCrLf &  RutaDocumentos & vbCrLf & LogoTipo & vbCrLf & RutaSysBases
     IniciarPrograma = False
  End If
End Sub

Public Sub ErrorDeImpresion()
  TextoError = "Error:(" & Err & ")" & vbCrLf _
             & "En la Impresora: " & Printer.DeviceName & vbCrLf _
             & "No pudo imprimir correctamente"
  MsgBox TextoError
  Printer.EndDoc
End Sub

Public Sub Keys_Especiales(KeyShift As Integer)
  ShiftDown = (KeyShift And vbShiftMask) > 0
  AltDown = (KeyShift And vbAltMask) > 0
  CtrlDown = (KeyShift And vbCtrlMask) > 0
End Sub

Public Sub PrinterPaint(PathDibujo As String, _
                        Xo As Single, _
                        Yo As Single, _
                        Xf As Single, _
                        Yf As Single)
  Xo = Redondear(Xo, 2)
  Yo = Redondear(Yo, 2)
  Xf = Redondear(Xf, 2)
  Yf = Redondear(Yf, 2)
  If Xo < 0 Then Xo = 0
  If Yo < 0 Then Yo = 0
  If Yo <= LimiteAlto And Len(PathDibujo) > 1 Then
     'MsgBox Dir(PathDibujo)
     If ((Dir(PathDibujo) <> "") And (Xf > 0) And (Yf > 0)) Then Printer.PaintPicture LoadPicture(PathDibujo), Xo, Yo, Xf, Yf
  End If
End Sub

Public Function ObtenerArchivo(Archivo As String) As String
Dim Io As Long
Dim Jo As Long
Dim Lo As Long
If Archivo <> Ninguno Then
   Archivo = RutaSistema & "\LOGOS\" & Archivo
   Archivo = Dir(Archivo)
   'MsgBox PathDibujo
   Lo = Len(Archivo)
   If Lo >= 1 Then
      Io = Lo: Jo = Lo
      Do While Io > 1
         If MidStrg(Archivo, Io, 1) = "." Then Jo = Io
         Io = Io - 1
      Loop
      ObtenerArchivo = MidStrg(Archivo, Io, Jo - Io)
   Else
      ObtenerArchivo = "DEFAULT"
   End If
Else
   ObtenerArchivo = "DEFAULT"
End If
End Function

Sub Ver_Image(NImage As Image, _
              Archivo As String)
    If Dir(Archivo) <> "" Then
       NImage.Picture = LoadPicture(Archivo)
    Else
       NImage.Picture = LoadPicture()
    End If
End Sub

Sub Escribir_Texto_Picture_Multiple(NPicture As PictureBox, _
                                    Xo As Single, _
                                    Yo As Single, _
                                    Color1 As Long, _
                                    Color2 As Long, _
                                    PictTexto As String)
Dim msg As String

Dim PColo As Single
Dim PFilo As Single

    If Len(PictTexto) > 1 Then
       PColo = Xo
       PFilo = Yo
       msg = ""
       AltoLetra = NPicture.TextHeight("H")
       For I = 1 To Len(PictTexto)
           If MidStrg(PictTexto, I, 1) = vbCr Or MidStrg(PictTexto, I, 1) = "^" Then
              If Color1 <> Color2 Then
                 NPicture.ForeColor = Color1
                 NPicture.CurrentX = PColo - 11
                 NPicture.CurrentY = PFilo - 10
                 NPicture.Print msg
              End If
              NPicture.ForeColor = Color2
              NPicture.CurrentX = PColo
              NPicture.CurrentY = PFilo
              NPicture.Print msg
              msg = ""
              PFilo = PFilo + AltoLetra
              I = I + 2
           End If
           msg = msg & MidStrg(PictTexto, I, 1)
       Next I
       If Color1 <> Color2 Then
          NPicture.ForeColor = Color1
          NPicture.CurrentX = PColo - 11
          NPicture.CurrentY = PFilo - 10
          NPicture.Print msg
       End If
       NPicture.ForeColor = Color2
       NPicture.CurrentX = PColo
       NPicture.CurrentY = PFilo
       NPicture.Print msg
    End If
End Sub
                         
Public Function Escribir_Texto_Picture_Ancho(NPicture As PictureBox, _
                                             PictTexto As String) As Single
Dim msg As String
Dim AnchoText As Single

    AnchoText = 0
    If Len(PictTexto) > 1 Then
       msg = ""
       AltoLetra = NPicture.TextHeight("H")
       For I = 1 To Len(PictTexto)
           If MidStrg(PictTexto, I, 1) = vbCr Or MidStrg(PictTexto, I, 1) = "^" Then
              If AnchoText < NPicture.TextWidth(msg) Then AnchoText = NPicture.TextWidth(msg)
              msg = ""
              I = I + 2
           End If
           msg = msg & MidStrg(PictTexto, I, 1)
       Next I
       If AnchoText < NPicture.TextWidth(msg) Then AnchoText = NPicture.TextWidth(msg)
    End If
    Escribir_Texto_Picture_Ancho = AnchoText
End Function
                         
Sub Escribir_Texto_Picture(NPicture As PictureBox, _
                           PictTexto As String)
Dim msg As String
Dim Web As String
Dim DatosEmpresa As String
Dim AnchoWeb As Single
Dim AltoWeb As Single
Dim Color1 As Long
Dim Color2 As Long
Dim Color3 As Long
Dim ColorC As Long
Dim CantEstadoDescripcion As Long
Dim Dias_Restantes As Long
Dim FinEjecucion As Boolean


  FinEjecucion = False
  
 'DatosEmpresa = "LICENCIA OTORGADA A LA ENTIDAD:" & vbCrLf
  DatosEmpresa = ""
  If UCaseStrg(Empresa) = UCaseStrg(NombreComercial) Then
     DatosEmpresa = DatosEmpresa _
                  & Empresa & " " & vbCrLf
  Else
     DatosEmpresa = DatosEmpresa _
                  & Empresa & " " & vbCrLf _
                  & NombreComercial & " " & vbCrLf
  End If
  DatosEmpresa = DatosEmpresa & "R.U.C. " & RUC & " " & vbCrLf
  If Len(Obligado_Conta) > 1 Then
     DatosEmpresa = DatosEmpresa & Obligado_Conta & " esta Obligado a LLevar Contabilidad " & vbCrLf
  Else
     DatosEmpresa = DatosEmpresa & "NO esta Obligado a LLevar Contabilidad " & vbCrLf
  End If
  If Len(ContEspec) > 1 Then DatosEmpresa = DatosEmpresa & ContEspec & "Contribuyente Especial " & vbCrLf
  
  Color1 = QBColor(Int((15 * Rnd) + 1)) 'Blanco
  Color2 = QBColor(Int((15 * Rnd) + 1))
  Color3 = QBColor(Int((15 * Rnd) + 1))
  If Color1 = Blanco Or Color1 = Amarillo Then Color1 = QBColor(Int((15 * Rnd) + 1))
  If Color1 = Color2 Then Color2 = QBColor(Int((15 * Rnd) + 1))
  If Color1 = Color3 Then Color3 = QBColor(Int((15 * Rnd) + 1))
  If Color2 = Color1 Then Color1 = QBColor(Int((15 * Rnd) + 1))
  If Color2 = Color3 Then Color3 = QBColor(Int((15 * Rnd) + 1))
  If Color1 = Blanco Or Color1 = Amarillo Then Color1 = QBColor(Int((15 * Rnd) + 1))
  
  NPicture.Font = TipoCourierNew  'TipoArialNarrow
  NPicture.FontBold = True
  NPicture.FontSize = 10
  NPicture.FontItalic = False
    
  PFil = 60
  PCol = 2650
  '1160
  NPicture.Line (0.1, 0.1)-(MDI_X_Max, 1400), Color2, BF
  NPicture.Line (50, 50)-(MDI_X_Max - 70, 1340), Blanco, BF
  If Len(LogoTipo) > 1 Then NPicture.PaintPicture LoadPicture(LogoTipo), 70, 70, 2500, 1050
  
  Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Azul, Azul, DatosEmpresa
    
  AnchoWeb = Escribir_Texto_Picture_Ancho(NPicture, DatosEmpresa)
  PFil = 60
  PCol = AnchoWeb + 3000
  DatosEmpresa = "Representante: " & NombreGerente & " " & vbCrLf _
               & "Direccion    : " & Direccion & " " & vbCrLf _
               & "Teléfono     : " & Telefono1 & " / FAX: " & FAX & " " & vbCrLf _
               & "Email Empresa: " & EmailEmpresa & " " & vbCrLf
  If Len(Lista_De_Correos(4).Correo_Electronico) > 1 Then
     DatosEmpresa = DatosEmpresa & "Email C.E.   : " & Lista_De_Correos(4).Correo_Electronico & " " & vbCrLf
  End If
  
  Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Azul, Azul, DatosEmpresa
  
  Color1 = QBColor(Int((15 * Rnd) + 1)) 'Blanco
  Color2 = QBColor(Int((15 * Rnd) + 1))
  Color3 = QBColor(Int((15 * Rnd) + 1))
  If Color1 = Blanco Then Color1 = QBColor(Int((15 * Rnd) + 1))
  If Color1 = Color2 Then Color2 = QBColor(Int((15 * Rnd) + 1))
  If Color1 = Color3 Then Color3 = QBColor(Int((15 * Rnd) + 1))
  If Color2 = Color1 Then Color1 = QBColor(Int((15 * Rnd) + 1))
  If Color2 = Color3 Then Color3 = QBColor(Int((15 * Rnd) + 1))
  If Color1 = Blanco Then Color1 = QBColor(Int((15 * Rnd) + 1))
  
  Web = "www.diskcoversystem.com"
  NPicture.FontBold = True
  NPicture.Font = TipoCourierNew
  RutaDestino = RutaSistema & "\LOGOS\diskcover_web.gif"
  NPicture.PaintPicture LoadPicture(RutaDestino), MDI_X_Max - 1500, 70, 1350, 500
  
  NPicture.FontSize = 12
  
 'Fecha de Renovacion del sistema
 '-------------------------------
  Dias_Restantes = CFechaLong(Fecha_CO) - CFechaLong(FechaSistema) + 1
  ColorC = Color_Dias_Restantes(Dias_Restantes)
  If Dias_Restantes > 0 Then
     msg = "Fecha Renovación del " & vbCrLf _
         & "Contrato: " & Fecha_CO & " " & vbCrLf _
         & Dias_Restantes & " Dia(s) Restante(s) "
  Else
     msg = "Contrato Vencido "
  End If
  msg = UCaseStrg(msg)
  AnchoWeb = Escribir_Texto_Picture_Ancho(NPicture, msg)
  PCol = MDI_X_Max - AnchoWeb - 1650
  PFil = 260
  Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Gris, ColorC, msg
 
 'Fecha de renovacion de Comprobantes Electronicos
 '------------------------------------------------
  Dias_Restantes = CFechaLong(Fecha_CE) - CFechaLong(FechaSistema) + 1
  ColorC = Color_Dias_Restantes(Dias_Restantes)
  If Dias_Restantes > 0 Then
     Select Case Ambiente
       Case "1": msg = "Comprobantes Electrónicos Ambiente de Prueba " & vbCrLf _
                     & "Fecha de Renovación: " & Fecha_CE & " " & vbCrLf _
                     & Dias_Restantes & " Dia(s) Restante(s) para su renovacion " & vbCrLf
       Case "2": msg = "Comprobantes Electrónicos Ambiente en Producción " & vbCrLf _
                     & "Fecha de Renovación: " & Fecha_CE & " " & vbCrLf _
                     & Dias_Restantes & " Dia(s) Restante(s) para su renovacion " & vbCrLf
       Case Else
                 msg = "Empresa sin Comprobantes Electrónicos " & vbCrLf
     End Select
  Else
     msg = "Empresa sin Comprobantes Electrónicos" & vbCrLf
  End If
  msg = UCaseStrg(msg)
  AnchoWeb = Escribir_Texto_Picture_Ancho(NPicture, msg)
  PCol = (MDI_X_Max - AnchoWeb) / 2 - 120
  PFil = 1500
  
  If Len(RutaCertificado) > 1 Then Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Gris, ColorC, msg
        
 'Fecha de renovacion del Certificado Electronico
 '-----------------------------------------------
  Dias_Restantes = CFechaLong(Fecha_P12) - CFechaLong(FechaSistema) + 1
  ColorC = Color_Dias_Restantes(Dias_Restantes)
  Select Case Ambiente
    Case "1", "2": msg = "Certificado Firma Electronica " & vbCrLf _
                       & "Fecha de Renovación: " & Fecha_P12 & " " & vbCrLf _
                       & Dias_Restantes & " Dia(s) Restante(s) para su renovacion " & vbCrLf
                   msg = UCaseStrg(msg)
                   AnchoWeb = Escribir_Texto_Picture_Ancho(NPicture, msg)
                   PCol = MDI_X_Max - AnchoWeb - 110
                   'PFil = 2800
                   PFil = 1500
                   If Len(RutaCertificado) > 1 Then Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Gris, ColorC, msg
  End Select
        
  msg = "Gerencia: gerencia@diskcoversystem.com " & vbCrLf _
      & "Teléfono: 099-965-4196 " & vbCrLf _
      & " " & vbCrLf & " " & vbCrLf _
      & Version_Sistema & " (" & MidStrg(RutaSistema, 1, 2) & ") " & vbCrLf
  PCol = 120
  PFil = Screen.Height - 3500
  NPicture.FontBold = False
  Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Color2, Blanco, msg

  msg = "Presidencia: prisma_net@hotmail.es " & vbCrLf _
      & "Teléfono   : 098-910-5300 " & vbCrLf _
      & " " & vbCrLf _
      & "Prisma Net : informacion@diskcoversystem.com " & vbCrLf _
      & "Teléfonos  : 099-467-8317 " & vbCrLf
  AnchoWeb = Escribir_Texto_Picture_Ancho(NPicture, msg)
  PCol = MDI_X_Max - AnchoWeb - 120
  'PFil = 4000
  PFil = Screen.Height - 3500
  NPicture.FontBold = False
  Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Color2, Blanco, msg

  AltoLetra = NPicture.TextHeight(MidStrg(msg, 1, 1))
  PCol = 100
  PFil = 1500
  Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Color1, Color2, PictTexto
   
  NPicture.FontSize = 14
  NPicture.FontItalic = False
  NPicture.FontBold = True
  NPicture.FontSize = 30
  NPicture.ForeColor = Blanco
  NPicture.FontName = TipoConsola
  AnchoWeb = NPicture.TextWidth(Web)
  PFil = MDI_Y_Max - 1600
  PCol = (MDI_X_Max - AnchoWeb) / 2
  NPicture.CurrentY = PFil
  NPicture.CurrentX = PCol  '150
  Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Color1, Color2, Web

''  NPicture.Print Web
''  NPicture.ForeColor = Color2 'Rojo
''  NPicture.CurrentY = PFil + 20
''  NPicture.CurrentX = PCol + 30
''  NPicture.Print Web
  
  msg = "Quito - Ecuador"
  NPicture.FontItalic = False
  NPicture.FontBold = True
  NPicture.FontSize = 16
  NPicture.ForeColor = Blanco
  NPicture.FontName = TipoConsola 'TipoTimes
  AnchoWeb = NPicture.TextWidth(msg)
  PFil = MDI_Y_Max - 700
  PCol = (MDI_X_Max - AnchoWeb) / 2
  NPicture.CurrentY = PFil
  NPicture.CurrentX = PCol  '150
  NPicture.Print msg
  NPicture.ForeColor = Color2 'Rojo
  NPicture.CurrentY = PFil + 20
  NPicture.CurrentX = PCol + 30
  NPicture.Print msg
  
  If EstadoEmpresa <> "OK" Then
     If Len(DescripcionEstado) = 1 Then DescripcionEstado = "EMPRESA SIN DEFINIR SU ESTADO"
     NPicture.FontName = TipoVerdana
     NPicture.FontItalic = False
     NPicture.FontSize = 32
     CantEstadoDescripcion = InStr(DescripcionEstado, vbCrLf)
     If CantEstadoDescripcion > 0 Then
        Cadena = MidStrg(DescripcionEstado, 1, CantEstadoDescripcion)
        
        AnchoWeb = NPicture.TextWidth(Cadena)
        PCol = (MDI_X_Max - AnchoWeb) / 2
        AnchoWeb = NPicture.TextHeight("H")
        PFil = (MDI_Y_Max - AnchoWeb) / 2
        NPicture.ForeColor = Blanco
        AltoWeb = PFil
        NPicture.CurrentY = PFil
        NPicture.CurrentX = PCol
        NPicture.Print Cadena
        
        PFil = PFil + 25
        PCol = PCol + 25
        NPicture.ForeColor = Rojo
        NPicture.CurrentY = PFil
        NPicture.CurrentX = PCol
        NPicture.Print Cadena
        
        Cadena = MidStrg(DescripcionEstado, CantEstadoDescripcion + 2, Len(DescripcionEstado))
        PFil = PFil + 650
     
        AnchoWeb = NPicture.TextWidth(Cadena)
        PCol = (MDI_X_Max - AnchoWeb) / 2
        AnchoWeb = NPicture.TextHeight("H")
        NPicture.ForeColor = Blanco
        AltoWeb = PFil
        NPicture.CurrentY = PFil
        NPicture.CurrentX = PCol
        NPicture.Print Cadena
        
        PFil = PFil + 25
        PCol = PCol + 25
        NPicture.ForeColor = Rojo
        NPicture.CurrentY = PFil
        NPicture.CurrentX = PCol
        NPicture.Print Cadena
     Else
        AnchoWeb = NPicture.TextWidth(DescripcionEstado)
        PCol = (MDI_X_Max - AnchoWeb) / 2
        AnchoWeb = NPicture.TextHeight("H")
        PFil = (MDI_Y_Max - AnchoWeb) / 2
        NPicture.ForeColor = Blanco
        AltoWeb = PFil
        NPicture.CurrentY = PFil
        NPicture.CurrentX = PCol
        NPicture.Print DescripcionEstado
        
        PFil = PFil + 25
        PCol = PCol + 25
        NPicture.ForeColor = Rojo
        NPicture.CurrentY = PFil
        NPicture.CurrentX = PCol
        NPicture.Print DescripcionEstado
     End If
     NPicture.FontSize = 12
     AnchoWeb = NPicture.TextWidth("ANCHO DE ALERTA")
     PCol = (MDI_X_Max - AnchoWeb) / 2
     PFil = PFil - 2800
     NPicture.Line (PCol - 40, PFil - 50)-(PCol + 2540, PFil + 2040), Azul, BF
     Select Case EstadoEmpresa
       Case "BLOQ", "MAS360": RutaDestino = RutaSistema & "\FORMATOS\BLOQUEO.jpg"
                              FinEjecucion = True
       Case "ONLY", "OKONLY": RutaDestino = RutaSistema & "\FORMATOS\ONLY.jpg"
       Case "DBMANT", "DBFULL": RutaDestino = RutaSistema & "\FORMATOS\Reparar.jpg"
       Case Else: RutaDestino = RutaSistema & "\FORMATOS\ADVERTENCIA.jpg"
     End Select
     NPicture.PaintPicture LoadPicture(RutaDestino), PCol, PFil, 2500, 2000
     ContadorEstados = ContadorEstados + 1
     If ContadorEstados > 6 Then
        Cadena = String(43, "_") & " " & vbCrLf _
               & " " & vbCrLf _
               & "M E N S A J E   D E   A D V E R T E N C I A" & vbCrLf _
               & "COMUNIQUECE AL DISTRIBUIDOR DEL SISTEMA, " & vbCrLf _
               & "O A NUESTRO CENTRO DE ATENCION AL CLIENTE: " & vbCrLf _
               & "TELEFONOS: 098-910-5300/099-965-4196 " & vbCrLf _
               & "EMAILS: prisma_net@hotmail.es " & vbCrLf _
               & "        soporte@diskcoversystem.com " & vbCrLf _
               & String(43, "_") & " "
        NPicture.Font = TipoCourierNew
        NPicture.FontSize = 18
        NPicture.FontBold = True
        AnchoWeb = Escribir_Texto_Picture_Ancho(NPicture, Cadena)
        PCol = (MDI_X_Max - AnchoWeb) / 2
        PFil = AltoWeb + 1000
        Escribir_Texto_Picture_Multiple NPicture, PCol, PFil, Blanco, Rojo, Cadena
        ContadorEstados = 0
     End If
     If FinEjecucion Then
        Cadena = String(64, "_") & " " & vbCrLf _
               & "M E N S A J E   D E   B L O Q U E O   D E F I N I T I V O" & vbCrLf _
               & " " & vbCrLf _
               & "COMUNIQUECE AL DISTRIBUIDOR DEL SISTEMA, O A NUESTRO " _
               & "CENTRO DE ATENCION AL CLIENTE: " & vbCrLf _
               & "TELEFONOS: 098-910-5300/099-965-4196 " & vbCrLf _
               & "EMAILS: prisma_net@hotmail.es/soporte@diskcoversystem.com " & vbCrLf _
               & String(64, "_") & " "
        MsgBox Cadena, vbCritical, ">>>>>  M E N S A J E   D E   A D V E R T E N C I A  <<<<<"
        End
     End If
  End If
  NPicture.FontBold = False
End Sub

Sub Ver_Picture(TipoEscala As Integer, _
                NPicture As PictureBox, _
                Archivo As String, _
                Optional Ancho As Single, _
                Optional alto As Single)
    If Dir(Archivo) <> "" Then
       NPicture.Picture = LoadPicture(Archivo)
    Else
       NPicture.Picture = LoadPicture()
    End If
End Sub

Sub Ver_Grafico_Form(NPicture As Form, _
                     Archivo As String)
    If Dir(Archivo) <> "" Then
       NPicture.Picture = LoadPicture(Archivo)
    Else
       NPicture.Picture = LoadPicture()
    End If
End Sub

Function IsFormLoaded(FormToCheck As Form) As Boolean
Dim y As Integer
    For y = 0 To Forms.Count - 1
        If Forms(y) Is FormToCheck Then
            IsFormLoaded = True
            Exit Function
        End If
    Next
    IsFormLoaded = False
End Function

Sub Ver_Grafico_FormPict(Optional CrearYa As Boolean)
Dim AdoEstado As ADODB.Recordset
Dim RecMySQL As ADODB.Recordset

Dim msg As String
Dim Texto_Ayuda As String
Dim RutaTexto As String
Dim RutaDestino1 As String
Dim RutaOrigen1 As String
Dim NombFilePict As String
Dim Email As String
Dim posPuntoComa As String

Dim RND_Files As Long

Dim Result As Integer
Dim Cont_Arch As Integer

Dim HalfWidth As Single
Dim HalfHeight As Single

   'MsgBox ContadorAyuda & vbCrLf & ContadorFondos
    If Not IP_PC.InterNet Then DescripcionEstado = "Su Ordenador no tiene conexion a internet, " & vbCrLf & "no podra autorizar Comprobantes Electronicos" Else DescripcionEstado = "No definido"
    
   'Obtenemos la fecha del Sistema
    FechaSistema = Format$(Date, FormatoFechas)
    
    If TiempoSistema = 0 Then TiempoSistema = Time
    If TiempoTarea = 0 Then TiempoTarea = Time
    Minutos = Time
    Segundos = Second(Minutos - TiempoSistema)
    Minutos = Minute(Minutos - TiempoSistema)
    MiTiempo = CSng(Format$(Minutos, "00") & "." & Format$(Segundos, "00"))
   'MsgBox Format$(Time, "HH:MM:SS") & vbCrLf & Format$(TiempoSistema, "HH:MM:SS") & vbCrLf & Minutos & vbCrLf & Segundos & vbCrLf & MiTiempo
   'If CrearYa Then MiTiempo = 6
    If MiTiempo >= 6 Then
      'MsgBox Minutos & vbCrLf & Segundos
      '-------------------------------------------------------------------------------
       TiempoSistema = Time
      '-------------------------------------------------------------------------------
       If Len(NumEmpresa) >= 3 And Len(NumModulo) > 1 And Len(CodigoUsuario) > 1 Then
          Cadena = ""
          
          NombFilePict = CodigoUsuario & NumEmpresa & NumModulo
         'MsgBox NombFilePict & vbCrLf & MDI_X_Max & " x " & MDI_Y_Max
          EstadoEmpresa = ""
          ComunicadoEntidad = ""
          MensajeEmpresa = ""
          IDEntidad = 0
          Fecha_CE = FechaSistema
          Result = 0
          PCActivo = True
          EstadoUsuario = True
         'Leemos el estado de la Empresa y su fecha de procesamiento
         'Fecha_CO,Fecha_CE,AgenteRetencion,MicroEmpresa,EstadoEmpresa,DescripcionEstado,NombreEntidad,RepresentanteLegal,MensajeEmpresa,ComunicadoEntidad
         '------------------------------------------------------------------------------------------------------------------------------------------------
          If IP_PC.InterNet Then
             Estado_Empresa_SP_MySQL
             If Len(ComunicadoEntidad) > 1 Or Len(MensajeEmpresa) > 1 Then
                Titulo = "COMUNICADO A LA ENTIDAD"
                Mensajes = "INFORMATIVO:" & vbCrLf _
                         & "-----------" & vbCrLf & vbCrLf _
                         & "Este es un mensaje automatico, enviado desde el Centro de " _
                         & "Control de Atención al Cliente, el informativo dice:" & vbCrLf & vbCrLf
                If Len(ComunicadoEntidad) > 1 Then Mensajes = Mensajes & UCaseStrg(ComunicadoEntidad) & vbCrLf & vbCrLf
                If Len(MensajeEmpresa) > 1 Then Mensajes = Mensajes & MensajeEmpresa & vbCrLf & vbCrLf
                Mensajes = Mensajes _
                         & "En caso de requerir atención personalizada por parte de un asesor de servicio " _
                         & "al cliente de DiskCover System, usted podrá solicitar ayuda mediante los canales de " _
                         & "atención al cliente oficiales que detallamos a continuación: " _
                         & "Telefonos: 099-965-4196/098-910-5300." & vbCrLf & vbCrLf _
                         & "Por la atención que se de al presente quedamos de usted." & vbCrLf & vbCrLf _
                         & "DESEA SEGUIR RECIBIENDO ESTE COMUNICADO"
                If BoxMensaje = vbNo Then
                  'Datos del destinatario de mails, Cod datos de privacidad
                   TMail.ListaMail = 5
                   TMail.Asunto = "Informativo enviado a Sr(a): " & RepresentanteLegal & ", representante de: " & NombreEntidad
                   TMail.Adjunto = ""
                   TMail.Credito_No = ""
                    
                  'Enviamos lista de mails
                   TMail.para = ""
                   Insertar_Mail TMail.para, CorreoDiskCover
                   Insertar_Mail TMail.para, EmailContador
                   Insertar_Mail TMail.para, EmailProcesos
                  'MsgBox "Lista: " & emails
'''                   Do While Len(Emails) > 3
'''                      posPuntoComa = InStr(Emails, ";")
'''                      Email = MidStrg(Emails, 1, posPuntoComa - 1)
'''                      TMail.para = Email
'''                     'MsgBox "Email: " & Email & vbCrLf & RutaXML
'''                      If EsUnEmail(TMail.para) Then FEnviarCorreos.Show 1
'''                      Emails = MidStrg(Emails, posPuntoComa + 1, Len(Emails))
'''                   Loop
'''
                   FEnviarCorreos.Show 1
                   
                   sSQL = "UPDATE entidad " _
                        & "SET Comunicado = '.' " _
                        & "WHERE ID_Empresa = " & IDEntidad & " "
                   Conectar_Ado_Execute_MySQL sSQL
                    
                   sSQL = "UPDATE lista_empresas " _
                        & "SET Mensaje = '.' " _
                        & "WHERE ID_Empresa = " & IDEntidad & " " _
                        & "AND RUC_CI_NIC = '" & RUC & "' " _
                        & "AND Item = '" & NumEmpresa & "' "
                   Conectar_Ado_Execute_MySQL sSQL
                End If
             End If
                            
             'Actualizamos el estado y fecha de comprobantes electronico de la empresa
              sSQL = "UPDATE Empresas " _
                   & "SET Estado = '" & EstadoEmpresa & "' " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Estado <> '" & EstadoEmpresa & "' "
              Ejecutar_SQL_SP sSQL
             
              sSQL = "UPDATE Empresas " _
                   & "SET Fecha_CE = '" & BuscarFecha(Fecha_CE) & "' " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Fecha_CE <> '" & BuscarFecha(Fecha_CE) & "' "
              Ejecutar_SQL_SP sSQL
              
             'Lista_Mensaje_SP_MySQL TextoFile
             '--------------------------------
              TextoFile = "DISKCOVER SYSTEM AHORA EN LAS NUBLES (CLOUD) "
          Else
              TextoFile = "DISKCOVER SYSTEM AHORA EN LAS NUBLES (CLOUD) " & vbCrLf & " " & vbCrLf & "COMUNIQUESE CON SU ADMINISTRADOR DE LA RED "
          End If
           
         'MsgBox DescripcionEstado
          
        'Colocamos el grafico si existe ya hecho en otras secciones sino colocamos el por default
         RutaDestino1 = RutaSistema & "\FONDOS\USUARIOS\" & NombFilePict & ".jpg"
         If Dir(RutaDestino1) = "" Then RutaDestino1 = RutaSistema & "\INICIO.jpg"
         'MDIFormulario.PictMDI.PaintPicture LoadPicture(RutaDestino1), 5, 1150, MDI_X_Max, MDI_Y_Max - 1000
         
        'Recuperamos el archivo de Fondo
         Cont_Arch = 1
         RND_Files = Int(((ContadorFondos - 1) * Rnd) + 1)
         
         If RND_Files < 1 Then RND_Files = 1
         If RND_Files > ContadorFondos Then RND_Files = ContadorFondos - 1
         RutaOrigen1 = Fondos_Pantalla(RND_Files)
         
        'De donde empieza a generarel grafico informativo de la Empresa
        '==============================================================
         MDIFormulario.PictMDI.Height = MDI_Y_Max
         MDIFormulario.PictMDI.Visible = True
         MDIFormulario.PictMDI.AutoRedraw = True
        'MsgBox MDI_X_Max & "x" & MDI_Y_Max & vbCrLf & Screen.width & "x" & Screen.Height
         
         MDIFormulario.PictMDI.PaintPicture LoadPicture(RutaOrigen1), 5, 1400, MDI_X_Max, MDI_Y_Max - 1000
         
        'Escribe el texto de los logos y los datos de la empresa
         Escribir_Texto_Picture MDIFormulario.PictMDI, TextoFile  'Msg
         
         RutaDestino1 = RutaSistema & "\FONDOS\USUARIOS\" & NombFilePict & ".jpg"
         SavePicture MDIFormulario.PictMDI.Image, RutaDestino1
         MDIFormulario.PictMDI.Visible = False
         MDIFormulario.Picture = LoadPicture(RutaDestino1)
    
         If Not PCActivo Then
            Cadena = NombreUsuario & vbCrLf & "Su Equipo se encuentra en LISTA NEGRA, ingreso no autorizado, comuniquese con el Administrador del Sistema"
            MsgBox UCaseStrg(Cadena), vbCritical, "ACCESO DEL PC DENEGADO"
            End
         End If
         If Not EstadoUsuario Then
            Cadena = NombreUsuario & vbCrLf & "Su ingreso no esta autorizado, comuniquese con el Administrador del Sistema"
            MsgBox UCaseStrg(Cadena), vbCritical, "ACCESO AL SISTEMA DENEGADO"
            End
         End If
       End If
    End If
End Sub

Public Function ClaveSupervisor() As Boolean
   TipoSuper = "Supervisor"
   IngClaves.Show 1
   ClaveSupervisor = ResultClaveSup
End Function

Public Function ClaveGerente() As Boolean
   TipoSuper = "Gerente"
   IngClaves.Show 1
   ClaveGerente = ResultClaveSup
End Function

Public Function ClaveContador() As Boolean
   TipoSuper = "Contador"
   IngClaves.Show 1
   ClaveContador = ResultClaveSup
End Function

Public Function ClaveAdministrador() As Boolean
   TipoSuper = "Administrador"
   IngClaves.Show 1
   ClaveAdministrador = ResultClaveSup
End Function

Public Function ClaveAuxiliar() As Boolean
   TipoSuper = "Auxiliar"
   IngClaves.Show 1
   ClaveAuxiliar = ResultClaveSup
End Function

Public Function ClaveReabrirPeriodo() As Boolean
   TipoSuper = "Reabrir"
   IngClaves.Show 1
   ClaveReabrirPeriodo = ResultClaveSup
End Function

Public Function BoxMensaje(Optional SegundoBoton As Boolean) As Integer
  RatonNormal
  TipoDeCaja = vbYesNo + vbQuestion
  If SegundoBoton Then TipoDeCaja = TipoDeCaja + vbDefaultButton2
  BoxMensaje = MsgBox(Mensajes, TipoDeCaja, UCaseStrg(Titulo))
End Function

Public Function ReadAdoCta(SQLs As String) As String
Dim Strgs As String
Dim NumCodigo As String
Dim AdoRegs As ADODB.Recordset
NumCodigo = Ninguno
If SQLs <> "" Then
   Set AdoRegs = New ADODB.Recordset
   AdoRegs.CursorType = adOpenStatic
   AdoRegs.CursorLocation = adUseClient
   Strgs = "SELECT * " _
         & "FROM Ctas_Proceso " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Detalle = '" & SQLs & "' "
   AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
   If AdoRegs.RecordCount > 0 Then NumCodigo = AdoRegs.fields("Codigo")
   AdoRegs.Close
 End If
 ReadAdoCta = NumCodigo
End Function

Public Function ReadSetDataNum(SQLs As String, _
                               ParaEmpresa As Boolean, _
                               Incrementar As Boolean) As Long
Dim AdoRegs As ADODB.Recordset
Dim Strgs As String
Dim NumEmpA As String
Dim MesComp As String
Dim NumCodigo As Long
Dim HoraDelSistema As Long
Dim Si_MesComp As Boolean
Dim NuevoNumero As Boolean

    NumCodigo = 0
    NuevoNumero = False
    If Len(FechaComp) < 10 Then FechaComp = FechaSistema
    If FechaComp = "00/00/0000" Then FechaComp = FechaSistema
    Si_MesComp = False
    If ParaEmpresa Then NumEmpA = NumEmpresa Else NumEmpA = "000"
    
    HoraDelSistema = Second(Time)
    HoraDelSistema = Int((HoraDelSistema * Rnd) + 1)
    If HoraDelSistema < 6 Then HoraDelSistema = 6
    Sleep HoraDelSistema
    
    If SQLs <> "" Then
       MesComp = ""
       If Len(FechaComp) >= 10 Then MesComp = Format$(Month(FechaComp), "00")
       If MesComp = "" Then MesComp = "01"
       
       If Num_Meses_CD And SQLs = "Diario" Then
          SQLs = MesComp & SQLs
          Si_MesComp = True
       End If
       If Num_Meses_CI And SQLs = "Ingresos" Then
          SQLs = MesComp & SQLs
          Si_MesComp = True
       End If
       If Num_Meses_CE And SQLs = "Egresos" Then
          SQLs = MesComp & SQLs
          Si_MesComp = True
       End If
       If Num_Meses_ND And SQLs = "NotaDebito" Then
          SQLs = MesComp & SQLs
          Si_MesComp = True
       End If
       If Num_Meses_NC And SQLs = "NotaCredito" Then
          SQLs = MesComp & SQLs
          Si_MesComp = True
       End If
          
       Strgs = "SELECT Numero, ID " _
             & "FROM Codigos " _
             & "WHERE Concepto = '" & SQLs & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Item = '" & NumEmpA & "' "
       Select_AdoDB AdoRegs, Strgs
      'MsgBox Strgs
       With AdoRegs
        If .RecordCount > 0 Then
            NumCodigo = .fields("Numero")
        Else
            NuevoNumero = True
            NumCodigo = 1
            If Num_Meses_CD And Si_MesComp Then NumCodigo = CLng(MesComp & "000001")
            If Num_Meses_CI And Si_MesComp Then NumCodigo = CLng(MesComp & "000001")
            If Num_Meses_CE And Si_MesComp Then NumCodigo = CLng(MesComp & "000001")
            If Num_Meses_ND And Si_MesComp Then NumCodigo = CLng(MesComp & "000001")
            If Num_Meses_NC And Si_MesComp Then NumCodigo = CLng(MesComp & "000001")
        End If
       End With
       AdoRegs.Close
    End If
    If NumCodigo > 0 Then
      'MsgBox NumCodigo
       If NuevoNumero Then
          Strgs = "INSERT INTO Codigos (Periodo,Item,Concepto,Numero) " _
                & "VALUES ('" & Periodo_Contable & "','" & NumEmpA & "','" & SQLs & "'," & NumCodigo & ") "
          Ejecutar_SQL_SP Strgs
       End If
       If Incrementar Then
          Strgs = "UPDATE Codigos " _
                & "SET Numero = Numero + 1 " _
                & "WHERE Concepto = '" & SQLs & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Item = '" & NumEmpA & "' "
          Ejecutar_SQL_SP Strgs
       End If
    End If
   'MsgBox NumCodigo
    ReadSetDataNum = NumCodigo
End Function

Public Function Leer_Codigo_Automatico() As String
Dim AdoRegs As ADODB.Recordset
Dim Strgs As String
Dim CodigoMayor As Long
Dim CodigoRUC As String

    Set AdoRegs = New ADODB.Recordset
    AdoRegs.CursorType = adOpenStatic
    AdoRegs.CursorLocation = adUseClient
   
   'Listado de Codigos para Facturacion de Bancos
    RatonReloj
    CodigoRUC = NumEmpresa & Format$(1, "00000")
    Strgs = "SELECT TOP 1 CI_RUC " _
          & "FROM Clientes " _
          & "WHERE ISNUMERIC(CI_RUC) <> 0 " _
          & "AND CI_RUC <= '" & NumEmpresa & "999999' " _
          & "AND MidStrg(CI_RUC,1,3) = '" & NumEmpresa & "' " _
          & "AND LEN(CI_RUC) <= 9 " _
          & "ORDER BY CI_RUC DESC "
    Strgs = CompilarSQL(Strgs)
    AdoRegs.open Strgs, AdoStrCnn, , , adCmdText
    CodigoMayor = 0
    With AdoRegs
     If .RecordCount > 0 Then
         CodigoMayor = Val(MidStrg(.fields("CI_RUC"), 4, Len(.fields("CI_RUC"))))
         CodigoRUC = NumEmpresa & Format$(CodigoMayor + 1, "00000")
     End If
    End With
    AdoRegs.Close
    RatonNormal
    Leer_Codigo_Automatico = CodigoRUC
End Function

Public Sub WriteSetDataNum(SQLs As String, _
                           ParaEmpresa As Boolean, _
                           Valor As Long)
Dim Strgs As String
Dim NumCodigo As Long
NumCodigo = 0
If SQLs <> "" Then
   Strgs = "SELECT * " _
         & "FROM Codigos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Concepto = '" & SQLs & "' "
   ConectarAdoRecordSet Strgs
   With AdoReg
    If .RecordCount > 0 Then
        Strgs = "UPDATE Codigos " _
              & "SET Numero = " & Valor & " " _
              & "WHERE Concepto = '" & SQLs & "' " _
              & "AND Item = '" & NumEmpresa & "' "
        Ejecutar_SQL_SP Strgs
    Else
        Strgs = "INSERT INTO Codigos(Item,Concepto,Numero) VALUES " _
             & "('" & NumEmpresa & "','" & SQLs & "'," & Valor & ") "
        Ejecutar_SQL_SP Strgs
    End If
    AdoReg.Close
   End With
 End If
End Sub

Public Sub SetProgBar(BarraProg As ProgressBar, _
                      Maximo As Long)
  BarraProg.Min = 0
  If BarraProg.Min >= Maximo Then Maximo = 100
  BarraProg.Max = Maximo
  BarraProg.value = BarraProg.Min
End Sub

Public Sub IncProgBar(BarraProg As ProgressBar)
  If BarraProg.value < BarraProg.Max Then BarraProg.value = BarraProg.value + 1
End Sub

Public Function Cambio_Letras(Num, Optional NumDecimales As Byte, Optional SinMayusc As Boolean, Optional SinSobre100 As Boolean) As String
Dim S1 As String
Dim NumLetra As String
Dim DigitoStr As String
Dim DigitoNum As Integer
Dim AuxDigitoNum As Integer
Dim PosDigito As String
Dim Digito As Long
  If NumDecimales > 0 Then
     S1 = Format$(Val(Num), "000,000,000,000." & String$(NumDecimales, "0"))
  Else
     S1 = Format$(Val(Num), "000,000,000,000")
  End If
  NumLetra = ""
  For Digito = 1 To 15
     DigitoStr = "": PosDigito = ""
     DigitoNum = Val(MidStrg(S1, Digito, 1))
    'Camibar a cadena:   123456789012345678
    'Formato del numero: 000,000,000,000.00
        Select Case Digito
          Case 1, 5, 9, 13
               If DigitoNum = 1 Then
                  For I = Digito + 1 To Digito + 2
                      If Val(MidStrg(S1, I, 1)) <> 0 Then PosDigito = "to"
                  Next I
               End If
          Case 3, 11
               For I = Digito To Digito - 2 Step -1
                  If Val(MidStrg(S1, I, 1)) <> 0 Then PosDigito = " mil"
               Next I
          Case 7
               For I = Digito To 1 Step -1
                  If Val(MidStrg(S1, I, 1)) <> 0 Then PosDigito = " millon"
               Next I
               If PosDigito <> "" Then
                  If DigitoNum <> 1 Then
                     PosDigito = PosDigito & "es "
                  Else
                     PosDigito = PosDigito & " "
                  End If
               End If
        End Select
     Select Case Digito
       Case 3, 7, 11, 15  'Unidades
            AuxDigitoNum = Val(MidStrg(S1, Digito - 1, 2))
            If 11 <= AuxDigitoNum And AuxDigitoNum <= 19 Then
               DigitoStr = ""
            Else
            If DigitoNum = 1 And Val(Num) <> 1 And Digito <> 15 Then
               DigitoStr = "un"
            Else
               DigitoStr = CambioUnidades(DigitoNum)
            End If
            
            End If
       Case 2, 6, 10, 14  'Decenas
            If DigitoNum <> 0 Then
            AuxDigitoNum = Val(MidStrg(S1, Digito, 2))
            If 11 <= AuxDigitoNum And AuxDigitoNum <= 19 Then
               DigitoStr = CambioDecenas1(AuxDigitoNum)
            Else
               AuxDigitoNum = Val(MidStrg(S1, Digito + 1, 1))
               If 1 <= AuxDigitoNum And AuxDigitoNum <= 9 Then
                  DigitoStr = CambioDecenas(DigitoNum, True)
               Else
                  DigitoStr = CambioDecenas(DigitoNum, False)
               End If
            End If
            End If
       Case 1, 5, 9, 13   'Centenas
            DigitoStr = CambioCentenas(DigitoNum)
     End Select
     If DigitoStr <> "" Or PosDigito <> "" Then
        NumLetra = NumLetra & DigitoStr & PosDigito & " "
     End If
  Next Digito
  If NumDecimales > 0 Then
     If SinSobre100 Then
        DigitoStr = TrimStrg(NumLetra) & ", " & MidStrg(S1, 16, 3)
     Else
        DigitoStr = TrimStrg(NumLetra) & MidStrg(S1, 16, 3) & "/100"
        DigitoStr = Replace(DigitoStr, ".", ", ")
     End If
  Else
     DigitoStr = TrimStrg(NumLetra)
  End If
  If SinMayusc = False Then DigitoStr = UCaseStrg(DigitoStr)
  Cambio_Letras = DigitoStr
End Function

Public Function Cambio_Letras_Decimales(Num, _
                                        Optional NumDecimales As Byte, _
                                        Optional SinMayusc As Boolean, _
                                        Optional Milesimas As Boolean) As String
Dim Num1 As Variant
Dim StrCLD As String
   Num = Redondear(Num, NumDecimales)
   Num1 = Num - Int(Num)
   If NumDecimales > 0 Then
      Num1 = Redondear(Num1 * CInt("1" & String$(NumDecimales, "0")))
   End If
   StrCLD = Cambio_Letras(Int(Num), 0, SinMayusc)
   If Num1 > 0 Then
      StrCLD = StrCLD & ", " & Cambio_Letras(Num1, 0, SinMayusc)
      If Milesimas Then StrCLD = StrCLD & " milésimas"
   End If
   If SinMayusc = False Then StrCLD = UCaseStrg(StrCLD)
   Cambio_Letras_Decimales = StrCLD
End Function

Public Function Cambio_Punto_x_Coma(Num As Variant) As String
Dim VEntera As Double
Dim VDecimal As Double
Dim VNumero As Double
Dim Cad_Coma As String
    If Num > 0 Then
       VNumero = Redondear_2Dec(Num)
       Cad_Coma = Replace(Format$(VNumero, "#.00"), ".", ",")
       'If Int(VNumero) < 10 Then Cad_Coma = "0" & Cad_Coma
    Else
       Cad_Coma = "-"
    End If
    Cambio_Punto_x_Coma = Cad_Coma
End Function

Public Sub UnidadSistema()
  RutaSistema = LeftStrg(CurDir$, 2) & "\SISTEMA"
  RutaSysBases = LeftStrg(CurDir$, 2) & "\SYSBASES"
  ChDir RutaSistema
End Sub

Public Function MesesLetras(Mes As Integer, Optional Mayuscula As Boolean) As String
Dim SMes As String
   SMes = ""
   Select Case Mes
     Case 1: SMes = "Enero"
     Case 2: SMes = "Febrero"
     Case 3: SMes = "Marzo"
     Case 4: SMes = "Abril"
     Case 5: SMes = "Mayo"
     Case 6: SMes = "Junio"
     Case 7: SMes = "Julio"
     Case 8: SMes = "Agosto"
     Case 9: SMes = "Septiembre"
     Case 10: SMes = "Octubre"
     Case 11: SMes = "Noviembre"
     Case 12: SMes = "Diciembre"
   End Select
   If Mayuscula Then SMes = UCaseStrg(SMes)
   MesesLetras = SMes
End Function

Public Function LetrasMeses(Mes As String) As Byte
Dim SMes As Byte
   If Mes = "" Then Mes = "ene"
   
   Select Case LCase$(MidStrg(Mes, 1, 3))
     Case "ene": SMes = 1
     Case "feb": SMes = 2
     Case "mar": SMes = 3
     Case "abr": SMes = 4
     Case "may": SMes = 5
     Case "jun": SMes = 6
     Case "jul": SMes = 7
     Case "ago": SMes = 8
     Case "sep": SMes = 9
     Case "oct": SMes = 10
     Case "nov": SMes = 11
     Case "dic": SMes = 12
     Case Else: SMes = 1
   End Select
   'MsgBox LCase$(MidStrg(Mes, 1, 3))
   LetrasMeses = SMes
End Function

Public Function DiasLetras(Mes As Integer) As String
Dim SMes As String
   SMes = ""
   Select Case Mes
     Case 1: SMes = "Domingo"
     Case 2: SMes = "Lunes"
     Case 3: SMes = "Martes"
     Case 4: SMes = "Miércoles"
     Case 5: SMes = "Jueves"
     Case 6: SMes = "Viernes"
     Case 7: SMes = "Sábado"
   End Select
   DiasLetras = SMes
End Function

Public Function BuscarFecha(FechaStr As String) As String
'dd/mm/yyyy
  If IsNumeric(FechaStr) Then
     If SQL_Server Then
        BuscarFecha = Format$(FechaSistema, "YYYYMMDD")
     Else
        BuscarFecha = Format$(FechaSistema, "YYYY-MM-DD")
     End If
     'MsgBox "Fecha Incorrecta"
  Else
     If SQL_Server Then
        BuscarFecha = Format$(FechaStr, "YYYYMMDD")
     Else
        BuscarFecha = Format$(FechaStr, "YYYY-MM-DD")
     End If
  End If
End Function

Public Function SinEspaciosIzq(Strg As String) As String
Dim Longitud As Long
Dim SinEspacios As String
  SinEspacios = "": Longitud = 1
  If Len(Strg) > 1 Then
  While (Longitud <= Len(Strg)) And (MidStrg(Strg, Longitud, 1) <> " ")
    SinEspacios = SinEspacios & MidStrg(Strg, Longitud, 1)
    Longitud = Longitud + 1
  Wend
  End If
  SinEspaciosIzq = TrimStrg(SinEspacios)
End Function

Public Function NumeroDeEspacios(Strg As String) As Integer
Dim NoEspacios As Integer
Dim ContLetras As Long
  NoEspacios = 0
  If Len(TrimStrg(Strg)) >= 1 Then
     For ContLetras = 1 To Len(TrimStrg(Strg))
         If MidStrg(TrimStrg(Strg), ContLetras, 1) = " " Then NoEspacios = NoEspacios + 1
     Next ContLetras
  End If
  NumeroDeEspacios = NoEspacios
End Function

Public Function SinEspaciosIzqNoBlancos(Strg As String, _
                                        NumeroBlanco As Integer) As String
Dim Longitud As Long
Dim SinEspacios As String
Dim Strg1 As String
  Strg1 = Strg
  For I = 1 To NumeroBlanco
   SinEspacios = "": Longitud = 1
   If Len(Strg1) > 1 Then
      While (Longitud <= Len(Strg1)) And (MidStrg(Strg1, Longitud, 1) <> " ")
         SinEspacios = SinEspacios & MidStrg(Strg1, Longitud, 1)
         Longitud = Longitud + 1
      Wend
      If SinEspacios = "" Then SinEspacios = "0"
      Strg1 = MidStrg(Strg1, Len(SinEspacios) + 1, Len(Strg1))
      J = 1
      While MidStrg(Strg1, J, 1) = " "
         J = J + 1
      Wend
      Strg1 = MidStrg(Strg1, J, Len(Strg1))
   End If
  Next I
  If SinEspacios = "" Then SinEspacios = "0"
  SinEspaciosIzqNoBlancos = SinEspacios
End Function

Public Function SinCodigoIzq(Strg As String) As String
Dim Longitud As Long
Dim SinEspacios As String
  SinEspacios = Strg: Longitud = 1
  If Len(Strg) > 1 Then
     While MidStrg(SinEspacios, Longitud, 1) <> " " And Longitud < Len(Strg)
           Longitud = Longitud + 1
     Wend
     While MidStrg(SinEspacios, Longitud, 1) = " " And Longitud < Len(Strg)
           Longitud = Longitud + 1
     Wend
     SinEspacios = MidStrg(Strg, Longitud, Len(Strg))
  End If
  If SinEspacios = "" Then SinEspacios = "0"
  SinCodigoIzq = TrimStrg(SinEspacios)
End Function

Public Function CodigoIzqGuion(Strg As String) As String
Dim Longitud As Long
Dim SinEspacios As String
  SinEspacios = Strg: Longitud = 1
  If Len(Strg) > 1 Then
     While MidStrg(SinEspacios, Longitud, 1) <> "-" And Longitud < Len(Strg)
           Longitud = Longitud + 1
     Wend
     SinEspacios = MidStrg(Strg, 1, Longitud - 1)
  End If
  If SinEspacios = "" Then SinEspacios = "0"
  CodigoIzqGuion = TrimStrg(SinEspacios)
End Function

Public Function ObtenerPalabra(Strg As String, _
                               NoWord As Byte) As String
Dim IndI As Long
Dim IndF As Long
Dim Ind2 As Long
Dim Longitud As Long
Dim InicStrg As Long
Dim FinStrg As Long
Dim Strg1 As String
Dim SinEspacios As String
Dim WordS() As String
ReDim WordS(NoWord + 1) As String
Strg1 = TrimStrg(Strg): SinEspacios = ""
If NoWord > 0 Then
   For Ind2 = 0 To NoWord
       WordS(Ind2) = ""
   Next Ind2
   Ind2 = 0
Volver:
   Longitud = Len(Strg1)
   IndI = 1: IndF = 1
   Do
     If MidStrg(Strg1, IndF, 1) = " " Then
        While MidStrg(Strg1, IndF, 1) = " " And IndF <= Longitud
              IndF = IndF + 1
        Wend
        If Ind2 <= NoWord Then
           IndF = IndF - 1
           Strg1 = MidStrg(Strg1, IndF + 1, Longitud)
           Ind2 = Ind2 + 1
           If Len(Strg1) > 0 Then GoTo Volver Else GoTo Salir
        Else
           IndF = Longitud + 1
        End If
     Else
        WordS(Ind2) = MidStrg(Strg1, IndI, IndF)
     End If
     IndF = IndF + 1
   Loop Until IndF > Longitud
Salir:
   SinEspacios = WordS(NoWord - 1)
End If
ObtenerPalabra = TrimStrg(SinEspacios)
End Function

Public Function ObtenerCampoTexto(Strg As String) As String
Dim Indice As Long
Dim Longitud As Long
Dim InicStrg As Long
Dim FinStrg As Long
Dim SinEspacios As String
InicStrg = 1: Indice = 1: Longitud = Len(Strg): SinEspacios = ""
If Longitud > 1 Then
   FinStrg = 1
   Do
     Indice = Indice + 1: FinStrg = FinStrg + 1
   Loop Until (Indice > Longitud) Or (MidStrg(Strg, Indice, 1) = "|")
   If MidStrg(Strg, InicStrg, FinStrg - 1) <> "" Then SinEspacios = MidStrg(Strg, InicStrg, FinStrg - 1)
   Indice = Indice + 1: InicStrg = Indice
End If
If SinEspacios = "" Then SinEspacios = Ninguno
ObtenerCampoTexto = TrimStrg(SinEspacios)
End Function

Public Function ObtenerNumPalabra(Strg As String, NoBlancos As Byte) As String
Dim Indice As Long
Dim Longitud As Long
Dim InicStrg As Long
Dim FinStrg As Long
Dim SinEspacios As String
InicStrg = 1: Indice = 1: Longitud = Len(Strg): SinEspacios = ""
If Longitud > 1 And NoBlancos > 0 Then
   FinStrg = 1
   For I = 1 To NoBlancos
     Do
       Indice = Indice + 1: FinStrg = FinStrg + 1
     Loop Until (Indice > Longitud) Or (MidStrg(Strg, Indice, 1) = " ")
     If MidStrg(Strg, InicStrg, FinStrg - 1) <> "" Then SinEspacios = MidStrg(Strg, 1, FinStrg)
     Indice = Indice + 1: InicStrg = Indice
   Next I
End If
If IsNull(SinEspacios) Or IsEmpty(SinEspacios) Then SinEspacios = Ninguno
ObtenerNumPalabra = TrimStrg(SinEspacios)
End Function

Public Function SinEspaciosDer(Strg As String) As String
Dim Longitud As Long
Dim SinEspacios As String
Dim UnaLetra As String
  SinEspacios = "": Longitud = Len(Strg)
  If Len(Strg) > 1 Then
     UnaLetra = MidStrg(Strg, Longitud, 1)
     Do While (Longitud >= 1) And (UnaLetra <> " ")
        SinEspacios = MidStrg(Strg, Longitud, 1) & SinEspacios
        UnaLetra = MidStrg(Strg, Longitud, 1)
        Longitud = Longitud - 1
     Loop
  End If
  If SinEspacios = "" Then SinEspacios = "0"
  SinEspaciosDer = TrimStrg(SinEspacios)
End Function

Public Function CodigoCorrecto(Strg As String) As String
Dim CadAux As String
    CadAux = Strg
    CadAux = Replace(CadAux, ":", " ")
    CadAux = Replace(CadAux, ".", " ")
    CadAux = Replace(CadAux, """", " ")
    CadAux = Replace(CadAux, "'", " ")
    CadAux = Replace(CadAux, "´", " ")
    If CadAux = "" Then CadAux = Ninguno
    CodigoCorrecto = CadAux
End Function

Public Function CodigosSinPuntos(Strg As String) As String
Dim CadAux As String
Dim Letra As String
  CadAux = ""
  For I = 1 To Len(Strg)
     Letra = MidStrg(Strg, I, 1)
     If Letra <> "." Then CadAux = CadAux & Letra
  Next I
  If CadAux = "" Then CadAux = Ninguno
  CodigosSinPuntos = CadAux
End Function

Public Function SinEspaciosCampo(Strg As String) As String
Dim Longitud As Long
Dim SinEspacios As String
Dim UnaLetra As String
  SinEspacios = Strg: Longitud = Len(Strg)
  If Longitud >= 1 Then
     Do
       Caracter = MidStrg(Strg, Longitud, 1)
       Longitud = Longitud - 1
     Loop Until Caracter <> " " Or Longitud <= 0
     Longitud = Longitud + 1
     If Longitud >= 1 Then SinEspacios = MidStrg(Strg, 1, Longitud)
  End If
  If SinEspacios = "" Then SinEspacios = "0"
  SinEspaciosCampo = TrimStrg(SinEspacios)
End Function

Public Function CompilarString(CadSQL, Optional LString As Long, Optional QuitarPuntos As Boolean) As String
Dim StrSQL As String
Dim StrSQLAux As String
Dim Indc As Long
 StrSQL = ""
 If LString > 0 Then CadSQL = MidStrg(CadSQL, 1, LString)
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "|" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> vbCr Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> vbLf Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "'" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "," Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
''' CadSQL = StrSQL
''' If Len(CadSQL) > 0 Then
'''    StrSQL = ""
'''    For Indc = 1 To Len(CadSQL)
'''     If MidStrg(CadSQL, Indc, 1) <> "-" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
'''    Next Indc
''' End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "$" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "#" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "&" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> Chr(39) Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
Next Indc
 End If
 StrSQLAux = StrSQL
 'MsgBox StrSQL
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "         ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "        ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "       ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "      ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "     ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "    ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "   ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "  ", " ", 1)
 If IsNull(StrSQL) Then StrSQL = "."
 If IsEmpty(StrSQL) Then StrSQL = "."
 StrSQL = TrimStrg(StrSQL)
 If Len(StrSQL) > 1 Then
    If MidStrg(StrSQL, 1, 1) = "." Then
       StrSQL = MidStrg(StrSQL, 2, Len(StrSQL) - 1)
    End If
 End If
 If Len(StrSQL) > 1 Then
    If MidStrg(StrSQL, Len(StrSQL), 1) = "." Then
       StrSQL = MidStrg(StrSQL, 1, Len(StrSQL) - 1)
    End If
 End If
 If StrSQL = "" Then StrSQL = Ninguno
 CompilarString = StrSQL
End Function

Public Function Compilar_Strg_Migracion(CadSQL, Optional LString As Long) As String
Dim StrSQL As String
Dim StrSQLAux As String
Dim Indc As Long
 If LString > 0 Then CadSQL = MidStrg(CadSQL, 1, LString)
 StrSQL = ""
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "|" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> vbCr Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> vbLf Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "'" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "," Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> "-" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 CadSQL = StrSQL
 If Len(CadSQL) > 0 Then
    StrSQL = ""
    For Indc = 1 To Len(CadSQL)
     If MidStrg(CadSQL, Indc, 1) <> Chr(39) Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
    Next Indc
 End If
 StrSQLAux = StrSQL
 'MsgBox StrSQL
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "         ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "        ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "       ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "      ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "     ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "    ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "   ", " ", 1)
 If Len(StrSQL) > 1 Then StrSQL = Replace(StrSQLAux, "  ", " ", 1)
 If IsNull(StrSQL) Then StrSQL = "."
 If IsEmpty(StrSQL) Then StrSQL = "."
 StrSQL = TrimStrg(StrSQL)
 If Len(StrSQL) > 1 Then
    If MidStrg(StrSQL, 1, 1) = "." Then
       StrSQL = MidStrg(StrSQL, 2, Len(StrSQL) - 1)
    End If
 End If
 If Len(StrSQL) > 1 Then
    If MidStrg(StrSQL, Len(StrSQL), 1) = "." Then
       StrSQL = MidStrg(StrSQL, 1, Len(StrSQL) - 1)
    End If
 End If
 If StrSQL = "" Then StrSQL = Ninguno
 Compilar_Strg_Migracion = StrSQL
End Function

Public Function Compilar_Texto_XML(CadXML As String) As String
Dim TextTemp   As String
   TextTemp = CadXML
   If Len(TextTemp) > 1 Then TextTemp = Replace(TextTemp, "ñ", "N", 1)
   If Len(TextTemp) > 1 Then TextTemp = Replace(TextTemp, "Ñ", "N", 1)
   Compilar_Texto_XML = TextTemp
End Function

Public Function CompilarRUC_CI(CadSQL As String) As String
Dim StrSQL As String
Dim Indc As Long
 If ContadorRUCCI <= 0 Then ContadorRUCCI = 1
'MsgBox CadSQL
 If (CadSQL = Ninguno) Or (Val(CadSQL) = 0) Then
    StrSQL = Format$(ContadorRUCCI, "00000000")
    ContadorRUCCI = ContadorRUCCI + 1
 ElseIf Len(CadSQL) < 8 Then
    StrSQL = CadSQL & String$(8 - Len(CadSQL), "0")
 Else
    StrSQL = ""
    If Len(CadSQL) > 0 Then
       StrSQL = ""
       For Indc = 1 To Len(CadSQL)
           If MidStrg(CadSQL, Indc, 1) <> "-" Then StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
       Next Indc
    End If
 End If
 If (Len(StrSQL) > 10) And (MidStrg(StrSQL, Len(StrSQL) - 2, 3) = "000") Then StrSQL = MidStrg(StrSQL, 1, Len(StrSQL) - 3)
 If Val(StrSQL) <= 0 Then
    StrSQL = Format$(ContadorRUCCI, "00000000")
    ContadorRUCCI = ContadorRUCCI + 1
 End If
 CompilarRUC_CI = StrSQL
End Function

Public Function SetearBlancos(Strg, _
                              LongStrg As Long, _
                              No_Blancos As Integer, _
                              EsNumero As Boolean, _
                              Optional ConLineas As Boolean, _
                              Optional Decimales As Boolean) As String
Dim Longitud As Long
Dim SinEspacios As String
Dim UnaLetra As String
  SinEspacios = ""
  If IsNull(Strg) Then Strg = ""
  If IsEmpty(Strg) Then Strg = ""
  Strg = CompilarString(Strg)
  Longitud = Len(Strg)
  If EsNumero Then
     If Decimales Then
        SinEspacios = Format$(Val(Strg), "##0.00")
     Else
        SinEspacios = Format$(Val(Strg), "##0")
     End If
     Longitud = Len(SinEspacios)
     If Longitud < LongStrg Then SinEspacios = Space(LongStrg - Longitud) & SinEspacios
  Else
     If LongStrg > 0 Then
        SinEspacios = Strg & Space(LongStrg)
        SinEspacios = MidStrg(SinEspacios, 1, LongStrg)
     Else
        SinEspacios = TrimStrg(Strg)
     End If
  End If
  If No_Blancos > 0 Then SinEspacios = SinEspacios & Space(No_Blancos)
  If ConLineas Then SinEspacios = SinEspacios & "|"
  'If ConLineas Then SinEspacios = SinEspacios & vbTab
  If SinEspacios = "" Then SinEspacios = " "
  'MsgBox SinEspacios
  SetearBlancos = SinEspacios
End Function

Public Function SetearCeros(Strg, _
                            LongStrg As Long, _
                            No_Blancos As Integer, _
                            EsNumero As Boolean, _
                            Optional ConLineas As Boolean, _
                            Optional Decimales As Boolean) As String
Dim Longitud As Integer
Dim SinEspacios As String
Dim UnaLetra As String
  SinEspacios = ""
  If IsNull(Strg) Then Strg = ""
  If IsEmpty(Strg) Then Strg = ""
  Strg = CompilarString(Strg)
  Longitud = Len(Strg)
  If EsNumero Then
     If Decimales Then
        SinEspacios = Format$(Strg, "##0.00")
     Else
        SinEspacios = Format$(Strg, "##0")
     End If
     Longitud = Len(SinEspacios)
     If Longitud < LongStrg Then SinEspacios = String$(LongStrg - Longitud, "0") & SinEspacios
  Else
     If LongStrg > 0 Then
        SinEspacios = Strg & Space(LongStrg)
        SinEspacios = MidStrg(SinEspacios, 1, LongStrg)
     Else
        SinEspacios = TrimStrg(Strg)
     End If
  End If
  If No_Blancos > 0 Then SinEspacios = SinEspacios & Space(No_Blancos)
  If ConLineas Then SinEspacios = SinEspacios & "|"
  If SinEspacios = "" Then SinEspacios = " "
  'MsgBox SinEspacios
  SetearCeros = SinEspacios
End Function

Public Function MigrarDatos(Strg As String, _
                            LStrg As Long, _
                            EsNum As Boolean, _
                            Digit As String, _
                            Optional Separador As String, _
                            Optional Decimales As Integer) As String
Dim Longitud As Integer
Dim SinEspacios As String
Dim UnaLetra As String
  SinEspacios = ""
  Strg = CompilarString(Strg, LStrg)
  If LStrg > 0 And Len(Strg) > 0 Then
     If EsNum Then
        If Decimales = -1 Then
           Strg = Format$(Val(Strg), "##0")
        Else
           Strg = Format$(Val(Strg), "##0.0000")
        End If
        For I = 1 To Len(Strg)
          If MidStrg(Strg, I, 1) <> "." Then SinEspacios = SinEspacios & MidStrg(Strg, I, 1)
        Next I
        Longitud = Len(SinEspacios)
        If LStrg < Longitud Then
           SinEspacios = MidStrg(SinEspacios, 1 + Longitud - LStrg, LStrg)
        Else
           SinEspacios = String$(LStrg - Longitud, Digit) & SinEspacios
        End If
     Else
        SinEspacios = Strg
        Longitud = Len(SinEspacios)
        SinEspacios = SinEspacios & String$(LStrg - Longitud, Digit)
     End If
  End If
  If Separador <> "" Then SinEspacios = SinEspacios & Separador
  If SinEspacios = "" Then SinEspacios = " "
  'MsgBox "[" & SinEspacios & "]"
  MigrarDatos = SinEspacios
End Function

Public Function FormatoFechaSpace(FechaStr As String) As String
Dim S1, S2, S3 As String
FechaStr = Format$(FechaStr, "dd/mm/yyyy")
S1 = LeftStrg(FechaStr, 2)
S2 = MidStrg(FechaStr, 4, 2)
S3 = RightStrg(FechaStr, 2)
FormatoFechaSpace = S1 & "    " & S2 & "    " & S3
End Function

Public Function FechaDia(FechaStr As String) As Integer
  FechaStr = Format$(FechaStr, "dd/mm/yyyy")
  FechaDia = Val(LeftStrg(FechaStr, 2))
End Function

Public Function FechaAnio(FechaStr As String) As Integer
  FechaStr = Format$(FechaStr, "dd/mm/yyyy")
  FechaAnio = Val(RightStrg(FechaStr, 4))
End Function

Public Function Convertir_Fecha(Mi_Fecha As String) As String
   Mi_Fecha = UCaseStrg(Replace(Mi_Fecha, "-", "/"))
   Mi_Fecha = Replace(Mi_Fecha, "JAN", "01")
   Mi_Fecha = Replace(Mi_Fecha, "FEB", "02")
   Mi_Fecha = Replace(Mi_Fecha, "MAR", "03")
   Mi_Fecha = Replace(Mi_Fecha, "APR", "04")
   Mi_Fecha = Replace(Mi_Fecha, "MAY", "05")
   Mi_Fecha = Replace(Mi_Fecha, "JUN", "06")
   Mi_Fecha = Replace(Mi_Fecha, "JUL", "07")
   Mi_Fecha = Replace(Mi_Fecha, "AUG", "08")
   Mi_Fecha = Replace(Mi_Fecha, "SEP", "09")
   Mi_Fecha = Replace(Mi_Fecha, "OCT", "10")
   Mi_Fecha = Replace(Mi_Fecha, "NOV", "11")
   Mi_Fecha = Replace(Mi_Fecha, "DEC", "12")
   Convertir_Fecha = Format$(CDate(Mi_Fecha), "dd/MM/yyyy")
End Function

Public Function FechaMes(FechaStr As String) As Integer
  FechaStr = Format$(FechaStr, "dd/mm/yyyy")
  FechaMes = Val(MidStrg(FechaStr, 4, 2))
End Function

Public Function PrimerDiaMes(FechaStr As String) As String
  If FechaStr = "00/00/0000" Then FechaStr = FechaSistema Else FechaStr = Format$(FechaStr, "dd/mm/yyyy")
  
  Mes = Month(FechaStr): Anio = Year(FechaStr)
  PrimerDiaMes = "01/" & Format$(Mes, "00") & "/" & Format$(Anio, "0000")
End Function

Public Function SiguienteMes(FechaStr As String, Optional Fin_Mes As Boolean) As String
  FechaStr = Format$(FechaStr, "dd/mm/yyyy")
  Dia = Day(FechaStr): Anio = Year(FechaStr)
  Mes = Month(FechaStr) + 1
  If Mes > 12 Then
     Mes = 1: Anio = Anio + 1
  End If
  Select Case Mes
    Case 1, 3, 5, 7, 8, 10, 12
         If (Dia > 31) Then Dia = 31
         If Fin_Mes Then Dia = 31
    Case 2
         If ((Anio Mod 4 <> 0) And (Dia > 28)) Then
            Dia = 28
            If Fin_Mes Then Dia = 28
         End If
         If ((Anio Mod 4 = 0) And (Dia > 29)) Then
            Dia = 29
            If Fin_Mes Then Dia = 29
         End If
    Case 4, 6, 9, 11
         If (Dia > 30) Then Dia = 30
         If Fin_Mes Then Dia = 30
    Case Else
         Dia = 30
  End Select
  SiguienteMes = Format$(Dia, "00") & "/" & Format$(Mes, "00") & "/" & Format$(Anio, "0000")
End Function

Public Function SiguienteAnio(FechaStr As String) As String
  FechaStr = Format$(FechaStr, "dd/mm/yyyy")
  Dia = Day(FechaStr)
  Mes = Month(FechaStr)
  Anio = Year(FechaStr) + 1
  If Mes = 2 Then
     If ((Anio Mod 4 <> 0) And (Dia > 28)) Then Dia = 28
     If ((Anio Mod 4 = 0) And (Dia > 29)) Then Dia = 29
  End If
  SiguienteAnio = Format$(Dia, "00") & "/" & Format$(Mes, "00") & "/" & Format$(Anio, "0000")
End Function

Public Function AnteriorAnio(FechaStr As String) As String
  FechaStr = Format$(FechaStr, "dd/mm/yyyy")
  Dia = Day(FechaStr)
  Mes = Month(FechaStr)
  Anio = Year(FechaStr) - 1
  If Mes = 2 Then
     If ((Anio Mod 4 <> 0) And (Dia > 28)) Then Dia = 28
     If ((Anio Mod 4 = 0) And (Dia > 29)) Then Dia = 29
  End If
  AnteriorAnio = Format$(Dia, "00") & "/" & Format$(Mes, "00") & "/" & Format$(Anio, "0000")
End Function

Public Function UltimoDiaMes(FechaStr As String) As String
Dim vDia As Integer
Dim Vmes As Integer
Dim Vanio As Integer
Dim vResultado As String
Dim vFechaStr As String

  vFechaStr = FechaStr
  If Not IsDate(vFechaStr) Then vFechaStr = FechaSistema
  If vFechaStr = "00/00/0000" Then vFechaStr = FechaSistema
  vFechaStr = Format$(vFechaStr, "dd/mm/yyyy")
  Vmes = Month(vFechaStr)
  Vanio = Year(vFechaStr)
  vDia = 31
  Select Case Vmes
    Case 4, 6, 9, 11
         vDia = 30
    Case 2
         vDia = 28
         If (Vanio Mod 4 = 0) Then vDia = 29
  End Select
  UltimoDiaMes = Format$(vDia, "00") & "/" & Format$(Vmes, "00") & "/" & Format$(Vanio, "0000")
End Function

Public Function MaximoDia(MM, ANO) As Integer
Dim vMaximoDia As Integer

    vMaximoDia = 31
    Select Case Val(MM)
      Case 4, 6, 9, 11
           vMaximoDia = 30
      Case 2
           vMaximoDia = 28
           If (ANO Mod 4 = 0) Then vMaximoDia = 29
    End Select
    MaximoDia = vMaximoDia
End Function

Public Function FormatoCodigo(S As String, N As Long) As String
Dim S1 As String
Dim S2 As String
Dim S3 As String
Dim S4 As String
I = 1: J = 1: K = 0
While (I <= Len(S)) And (K = 0)
  If MidStrg(S, I, 1) = " " Then K = I
  I = I + 1
Wend
S1 = "": S2 = ""
If K > 0 Then
   S2 = MidStrg(S, K + 1, 1)
   S1 = LeftStrg(S, 2)
Else
   S1 = LeftStrg(S, 3)
End If
S3 = UCaseStrg(S1 & S2 & Format$(N, "00"))
S4 = ""
For I = 1 To Len(S3)
    If MidStrg(S3, I, 1) = " " Then
       S4 = S4 & "."
    Else
       S4 = S4 & MidStrg(S3, I, 1)
    End If
Next I
FormatoCodigo = NumEmpresa & S4
End Function

Public Function FormatoCodigo6(S As String, N As Long) As String
Dim S1 As String
Dim S2 As String
Dim S3 As String
Dim S4 As String
I = 1: J = 1: K = 0
While (I <= Len(S)) And (K = 0)
  If MidStrg(S, I, 1) = " " Then K = I
  I = I + 1
Wend
S1 = "": S2 = ""
If K > 0 Then
   S2 = MidStrg(S, K + 1, 1)
   S1 = LeftStrg(S, 1)
Else
   S1 = LeftStrg(S, 2)
End If
S3 = UCaseStrg(S1 & S2 & Format$(N, "0000"))
S4 = ""
For I = 1 To Len(S3)
    If MidStrg(S3, I, 1) = " " Then
       S4 = S4 & "."
    Else
       S4 = S4 & MidStrg(S3, I, 1)
    End If
Next I
FormatoCodigo6 = S4
End Function

Public Function FechaStrg(Fechas As String) As String
Dim dd, MM, AA As String
Fechas = Format$(Fechas, FormatoFechas)
  dd = Day(Fechas)
  MM = Month(Fechas)
  AA = Year(Fechas)
  If (Val(AA) < 1) Then AA = Anio
  If ((Val(MM) < 1) Or (Val(MM) > 12)) Then MM = Mes
  If (Val(dd) < 1) Or (Val(dd) > MaximoDia(MM, AA)) Then dd = MaximoDia(MM, AA)
  FechaStrg = dd & " de " & MesesLetras(Val(MM)) & " del " & AA
End Function

Public Function FechaStrgCiudad(Fechas As String) As String
Dim dd, MM, AA As String
 If NombreCiudad <> "" Then
    Fechas = Format$(Fechas, FormatoFechas)
    dd = Day(Fechas)
    MM = Month(Fechas)
    AA = Format$(Year(Fechas), "0000")
    If (Val(AA) < 1) Then AA = Anio
    If ((Val(MM) < 1) Or (Val(MM) > 12)) Then MM = Mes
    If (Val(dd) < 1) Or (Val(dd) > MaximoDia(MM, AA)) Then dd = MaximoDia(MM, AA)
    
    FechaStrgCiudad = ULCase(NombreCiudad) & ", " & dd & " de " & LCase$(MesesLetras(Val(MM))) & " del " & AA
 Else
    FechaStrgCiudad = " "
 End If
End Function

Public Function FechaDeCierre(Fechas As String) As String
Dim FechaC As String
Dim FechaCar As String
  FechaC = FechaStrgCorta(Fechas)
  FechaCar = ""
  For I = 1 To Len(FechaC)
      If MidStrg(FechaC, I, 1) = "/" Then
         FechaCar = FechaCar & "_"
      Else
         FechaCar = FechaCar & MidStrg(FechaC, I, 1)
      End If
  Next I
  FechaC = FechaCar
  FechaCar = ""
  For I = 1 To 12
      If I = 9 Then
         FechaCar = FechaCar & "." & MidStrg(FechaC, I, 1)
      Else
         FechaCar = FechaCar & MidStrg(FechaC, I, 1)
      End If
  Next I
  FechaDeCierre = FechaCar
End Function

Public Function FechaStrgCorta(Fechas As String) As String
Dim dd, MM, AA As String
  If Fechas = "00/00/0000" Then Fechas = FechaSistema
  Fechas = Format$(Fechas, FormatoFechas)
  dd = Format$(Day(Fechas), "00")
  MM = Format$(Month(Fechas), "00")
  AA = Format$(Year(Fechas), "0000")
  If (Val(AA) < 1) Then AA = Anio
  If ((Val(MM) < 1) Or (Val(MM) > 12)) Then MM = Mes
  If (Val(dd) < 1) Or (Val(dd) > MaximoDia(MM, AA)) Then dd = MaximoDia(MM, AA)
  FechaStrgCorta = dd & "/" & MidStrg(UCaseStrg(MesesLetras(Val(MM))), 1, 3) & "/" & AA
End Function

Public Function CentrarTexto(Texto As String, Optional LimAncho As Single) As Single
   If LimAncho > 0 Then
      CentrarTexto = (LimAncho / 2) - (Printer.TextWidth(Texto) / 2)
   Else
      CentrarTexto = (LimiteAncho / 2) - (Printer.TextWidth(Texto) / 2)
   End If
End Function

Public Sub PrinterCentrarTexto(LimAncho As Single, PosLineaY As Single, Texto As String)
Dim CentrarTextoW As Single
   If LimAncho > 0 Then
      CentrarTextoW = (LimAncho / 2) - (Printer.TextWidth(Texto) / 2)
   Else
      CentrarTextoW = (LimiteAncho / 2) - (Printer.TextWidth(Texto) / 2)
   End If
   If CentrarTextoW < 0 Then CentrarTextoW = 0.01
   If CentrarTextoW > 0 And PosLineaY > 0 Then
      Printer.CurrentX = CentrarTextoW
      Printer.CurrentY = PosLineaY
      Printer.Print Texto
   End If
End Sub

Public Function CentrarTextoEncab(Texto As String, X0, x1) As Single
   CentrarTextoEncab = (X0 / 2) + (x1 / 2) - (Printer.TextWidth(Texto) / 2)
End Function

Public Function CentrarTextoMargen(Texto As String) As Single
   CentrarTextoMargen = (MargenDer / 2) - (Printer.TextWidth(Texto) / 2)
End Function

Public Function TextoDerecha(Texto As String) As Single
   TextoDerecha = LimiteAncho - Printer.TextWidth(Texto)
End Function

Public Function VariableWidth(Variable) As Single
Dim DistVar As Single
DistVar = 0
StrgFormatoVariable = " "
Select Case VarType(Variable)
  Case vbBoolean: If Variable Then StrgFormatoVariable = "X"
  Case vbByte, vbInteger, vbLong
       StrgFormatoVariable = Str(Variable)
  Case vbSingle
       StrgFormatoVariable = Format$(Variable, "##0.00%")
  Case vbDouble, vbCurrency
       StrgFormatoVariable = Format$(Variable, "#,##0.00")
  Case vbString:  StrgFormatoVariable = Variable
  Case Else:      StrgFormatoVariable = Variable
End Select
StrgAncho = Ancho_Tipo_Variable(Variable)
AltoLetra = Printer.TextHeight(StrgAncho)
LongNumero = Printer.TextWidth(StrgAncho)
LongVariable = Printer.TextWidth(StrgFormatoVariable)
Select Case VarType(Variable)
  Case vbByte, vbDouble, vbInteger, vbLong, vbSingle, vbCurrency
       DistVar = LongNumero + 0.05 - LongVariable
       If DistVar <= 0 Then DistVar = 0.05
       If Printer.FontBold Then DistVar = DistVar - 0.04
       'If Val(StrgFormatoVariable) = 0 Then StrgFormatoVariable = "0"
  Case Else
       DistVar = 0
End Select
If DistVar <= 0 Then DistVar = 0.05
VariableWidth = DistVar
End Function

Function Ancho_Tipo_Variable(Variable) As String
Dim ch As String
    Select Case VarType(Variable)
      Case vbBoolean:  ch = CadBoolean
      Case vbByte:     ch = CadByte
      Case vbInteger:  ch = CadInteger
      Case vbLong:     ch = CadLong
      Case vbSingle:   ch = CadSingle
      Case vbDouble:   ch = CadDouble
      Case vbCurrency: ch = CadCurrency
      Case vbString:   ch = String$(Len(Variable), "X") & " "
      Case Else:       ch = String$(15, "X") & " "
    End Select
    If ch = "" Then ch = " "
    If ch = "." Then ch = " "
    Ancho_Tipo_Variable = ch
End Function

Public Function FormatoTipoCampo(AdoTipo As ADODB.Field) As String
Dim ch As String
Dim IdxCampo As Integer
Dim FormatoDeDecimales As String
  With AdoTipo
   'MsgBox .Type & vbCrLf & .Name & vbCrLf & .value
    Select Case .Type
      Case TadBoolean
           If AdoTipo Then ch = "X" Else ch = " "
      Case TadDate, TadDate1
           ch = Format$(AdoTipo, FormatoFechas)
      Case TadByte, TadInteger, TadLong
           If AdoTipo <> 0 Then ch = Format$(AdoTipo, "##0") Else ch = "0"
      Case TadDouble, TadCurrency
           FormatoDeDecimales = "#,##0.00"
           For IdxCampo = 0 To UBound(Vect_Dec) - 1
               If .Name = Vect_Dec(IdxCampo).Campo Then
                  'MsgBox .Name & ": " & IdxCampo & " ====>>>> " & Vect_Dec(IdxCampo).Campo & vbCrLf & Vect_Dec(IdxCampo).CantDec
                   FormatoDeDecimales = "#,##0." & String$(Vect_Dec(IdxCampo).CantDec, "0")
               End If
           Next IdxCampo
           If AdoTipo <> 0 Then ch = Format$(AdoTipo, FormatoDeDecimales) Else ch = "0.00"
           ' MsgBox .Name & vbCrLf & AdoTipo & vbCrLf & ch
      Case TadSingle
           If AdoTipo <> 0 Then ch = Format$(AdoTipo, "##0.00%") Else ch = "0.00%"
      Case Else
           If IsNull(AdoTipo) Then ch = "" Else ch = AdoTipo
    End Select
  End With
  FormatoTipoCampo = ch
End Function

Public Function FormatoTipoVariable(Variable) As String
Dim ch As String
    Select Case VarType(Variable)
        Case vbBoolean
             If Variable Then ch = "X" Else ch = " "
        Case vbDate
             ch = Format$(Variable, FormatoFechas)
        Case vbByte, vbInteger, vbLong
            ch = Format$(Variable, "##0")
        Case vbSingle
            ch = Format$(Variable, "##0.00%")
        Case vbDouble, vbCurrency
            ch = Format$(Variable, "#,##0.00")
        Case Else
            ch = Variable
    End Select
    If ch = "" Then ch = " "
    ch = ch & " "
    FormatoTipoVariable = ch
End Function

Public Sub PrinterVariables(X1o As Single, Y1o As Single, Variable, Optional ImpLineaCeros As Boolean)
'If Y1o <= LimiteAlto Then
   If ((X1o > 0) And (Y1o > 0)) Then
      Distancia = VariableWidth(Variable)
      If StrgFormatoVariable = Ninguno Then
         StrgFormatoVariable = " "
      ElseIf StrgFormatoVariable = "0" Or StrgFormatoVariable = "0.00" Then
        If Not ImpLineaCeros Then StrgFormatoVariable = " "
      End If
      If Y1o <= LimiteAlto Then
         Printer.CurrentX = X1o + CSng(Distancia)
         Printer.CurrentY = Y1o
         Printer.Print StrgFormatoVariable
      End If
     'MsgBox Printer.CurrentX & vbCrLf & Printer.CurrentY & vbCrLf & StrgFormatoVariable
   End If
'End If
End Sub

Public Sub PrinterVariableTexto(Xo As Single, Yo As Single, Texto As String, Variable)
Dim CantStrg As Single
'If Yo <= LimiteAlto Then
If ((Xo > 0) And (Yo > 0)) Then
   CantStrg = Printer.TextWidth(Texto) + 0.1
   Distancia = VariableWidth(Variable)
   'If Yo <= LimiteAlto Then
      Printer.CurrentX = Xo
      Printer.CurrentY = Yo
      If Texto = Ninguno Then Texto = " "
      If Texto = "0" Then Texto = " "
      Printer.Print Texto
      Printer.CurrentY = Yo
      Printer.CurrentX = Xo + CantStrg + Distancia
      If StrgFormatoVariable = Ninguno Then StrgFormatoVariable = " "
      If StrgFormatoVariable = "0" Then StrgFormatoVariable = " "
      If StrgFormatoVariable = "0.00" Then StrgFormatoVariable = " "
      Printer.Print StrgFormatoVariable
   'End If
End If
'End If
End Sub

Public Sub PrinterTexto(Xo As Single, _
                        Yo As Single, _
                        Texto As String, _
                        Optional JustDer As Boolean, _
                        Optional anchoTexto As Single)
'If Yo <= LimiteAlto Then
If ((Xo > 0) And (Yo > 0) And (Texto <> "")) Then
   RatonReloj
   'If Yo < (LimiteAlto - 0.35) Then
      LimpiarLineaTexto Xo, Yo, Texto
      If JustDer Then
         Printer.CurrentX = Xo + anchoTexto - Printer.TextWidth(Texto) + 0.1
      Else
         Printer.CurrentX = Xo + 0.1
      End If
      If Texto = Ninguno Then Texto = " "
      'If Texto = "0" Then Texto = " "
      Printer.CurrentY = Yo
      Printer.Print Texto
   'End If
   RatonNormal
End If
'End If
End Sub

Public Function CambioUnidades(Digitos As Integer) As String
Dim ResultDig As String
ResultDig = ""
  Select Case Digitos
    Case 1: ResultDig = "uno"
    Case 2: ResultDig = "dos"
    Case 3: ResultDig = "tres"
    Case 4: ResultDig = "cuatro"
    Case 5: ResultDig = "cinco"
    Case 6: ResultDig = "seis"
    Case 7: ResultDig = "siete"
    Case 8: ResultDig = "ocho"
    Case 9: ResultDig = "nueve"
  End Select
  CambioUnidades = ResultDig
End Function

Public Function CambioDecenas(Digitos As Integer, OpcY As Boolean) As String
Dim ResultDig As String
ResultDig = ""
  Select Case Digitos * 10
    Case 10: ResultDig = "diez"
    Case 20: ResultDig = "veinte"
    Case 30: ResultDig = "treinta"
    Case 40: ResultDig = "cuarenta"
    Case 50: ResultDig = "cincuenta"
    Case 60: ResultDig = "sesenta"
    Case 70: ResultDig = "setenta"
    Case 80: ResultDig = "ochenta"
    Case 90: ResultDig = "noventa"
  End Select
  If OpcY Then ResultDig = ResultDig & " y"
  CambioDecenas = ResultDig
End Function

Public Function CambioDecenas1(Digitos As Integer) As String
Dim ResultDig As String
ResultDig = ""
  Select Case Digitos
    Case 11: ResultDig = "once"
    Case 12: ResultDig = "doce"
    Case 13: ResultDig = "trece"
    Case 14: ResultDig = "catorce"
    Case 15: ResultDig = "quince"
    Case 16: ResultDig = "dieciseis"
    Case 17: ResultDig = "diecisiete"
    Case 18: ResultDig = "dieciocho"
    Case 19: ResultDig = "diecinueve"
  End Select
  CambioDecenas1 = ResultDig
End Function

Public Function CambioCentenas(Digitos As Integer) As String
Dim ResultDig As String
ResultDig = ""
  Select Case Digitos * 100
    Case 100: ResultDig = "cien"
    Case 200: ResultDig = "doscientos"
    Case 300: ResultDig = "trescientos"
    Case 400: ResultDig = "cuatrocientos"
    Case 500: ResultDig = "quinientos"
    Case 600: ResultDig = "seiscientos"
    Case 700: ResultDig = "setecientos"
    Case 800: ResultDig = "ochocientos"
    Case 900: ResultDig = "novecientos"
  End Select
  CambioCentenas = ResultDig
End Function

Public Sub PrinterNum(Xo As Single, Yo As Single, Numero As Currency)
Dim Numero_Letras As String
'If Yo < LimiteAlto Then
If ((Xo > 0) And (Yo > 0)) Then
   Numero_Letras = Cambio_Letras(Numero, 2)
   PrinterLineas Xo, Yo, Numero_Letras, 15.5
End If
'End If
End Sub

Public Sub PrinterNumCheque(Xo As Single, Yo As Single, AnchoCheque As Single, ValorChq As Currency)
Dim Numero_Letras As String
Dim I As Long
Dim Indc As Byte
If ((Xo > 0) And (Yo > 0)) Then
   Numero_Letras = UCaseStrg(Cambio_Letras(ValorChq, 2) & "." & String$(250, " "))
   IR = Redondear(AnchoCheque * 2, 4)
   JR = Redondear(Printer.TextWidth(Numero_Letras), 4)
   While JR >= IR
     Numero_Letras = MidStrg(Numero_Letras, 1, Len(Numero_Letras) - 1)
     JR = Printer.TextWidth(Numero_Letras)
   Wend
   IR = AnchoCheque
   JR = 0: I = 0
   While (JR < IR) And (I <= Len(Numero_Letras))
     I = I + 1
     Cadena = MidStrg(Numero_Letras, 1, I)
     'MsgBox Numero_Letras & vbCrLf & Cadena & vbCrLf & JR & " - " & IR
     JR = Redondear(Printer.TextWidth(Cadena), 4)
   Wend
   Cadena = MidStrg(Numero_Letras, 1, I)
   'If Yo < LimiteAlto Then
      Printer.CurrentX = Xo
      Printer.CurrentY = Yo
      Printer.Print Cadena
   'End If
   Numero_Letras = MidStrg(Numero_Letras, I + 1, Len(Numero_Letras))
   'If (Yo + 0.6) < LimiteAlto Then
      Printer.CurrentX = Xo - 1.5: Printer.CurrentY = Yo + 0.6
      Printer.Print Numero_Letras
   'End If
End If
End Sub

Public Function FTextWidth(Texto As String) As Single
  FTextWidth = Len(Texto) * 100
End Function

Public Sub Imprimir_Linea_H(Yo As Single, _
                            Xo As Single, _
                            Xf As Single, _
                            Optional ColorLinea As Long, _
                            Optional DobleLinea As Boolean)
   'MsgBox ColorLinea
   If Xf > AnchoPapel Then Xf = AnchoPapel
   If (0 < Yo) And (Yo <= SetPapelLargo) Then
      Printer.Line (Xo, Yo)-(Xf, Yo), ColorLinea
      If DobleLinea Then
         Yo = Yo + 0.025
         Printer.Line (Xo, Yo)-(Xf, Yo), ColorLinea
         Yo = Yo + 0.025
         Printer.Line (Xo, Yo)-(Xf, Yo), ColorLinea
      End If
   End If
End Sub

Public Sub Imprimir_Linea_V(Xo As Single, _
                            Yo As Single, _
                            Yf As Single, _
                            Optional ColorLinea As Long)
   If Yf > LimiteAlto Then Yf = LimiteAlto
   If Yo > Yf Then Yo = Yf
   If (0 < Xo) And (Xo <= AnchoPapel) Then Printer.Line (Xo, Yo - 0.05)-(Xo, Yf + 0.05), ColorLinea
End Sub

Public Sub Imprimir_Recibo_Caja(TRecibo As Tipo_Recibo)
Dim Ini_X As Double
Dim Ini_Y As Double
Dim Copia As Byte
Dim tipoDeLetra As String
Dim printLinea As String
Dim PrintCar As String
  
    RatonNormal
    Copia = 2
    Titulo = "TIPO DE IMPRESION"
    Mensajes = "Imprimir Recibo de Caja con copia?"
    If BoxMensaje = vbYes Then Copia = 0
    RatonReloj
    tipoDeLetra = TipoCourier 'TipoTimes
   'Geneeramos el documento
    tPrint.TipoImpresion = Es_PDF
    tPrint.NombreArchivo = "Recibo_No_" & Format$(TRecibo.Recibo_No, "000000000")
    tPrint.TituloArchivo = "Recibo de Caja"
    tPrint.TipoLetra = tipoDeLetra
    tPrint.OrientacionPagina = 1
    tPrint.PaginaA4 = True
    tPrint.EsCampoCorto = False
    tPrint.VerDocumento = True
    
    Set cPrint = New cImpresion
    cPrint.iniciaImpresion
Imprimir_Copia:
    'cPrint.printCuadroLinea 2, 2, 2, 2, Negro, "B"
    If Copia = 0 Then Ini_X = 1.3 Else Ini_X = 11.5
    cPrint.printImagen LogoTipo, Ini_X + 0.1, 1, 2.5, 1
    PosLinea = 1
    cPrint.letraTipo tipoDeLetra, 8
    'cPrint.colorDeLetra = Negro
    PosLinea = cPrint.printTextoMultiple(Ini_X + 2.7, PosLinea, UCaseStrg(RazonSocial), 6.8)
    PosLinea = PosLinea + 0.35
    PosLinea = cPrint.printTextoMultiple(Ini_X + 2.7, PosLinea, UCaseStrg(NombreComercial), 6.8)
    PosLinea = PosLinea + 0.35
    cPrint.printTexto Ini_X + 2.7, PosLinea, "R.U.C.: " & RUC
    PosLinea = PosLinea + 0.4
    cPrint.letraTipo tipoDeLetra, 6
    PosLinea = cPrint.printTextoMultiple(Ini_X, PosLinea, "DIRECCION: " & Direccion, 9.5)
    PosLinea = PosLinea + 0.3
    If UCaseStrg(Direccion) <> UCaseStrg(DireccionEstab) Then
       PosLinea = cPrint.printTextoMultiple(Ini_X, PosLinea, "SUCURSAL: " & DireccionEstab, 9.5)
       'cPrint.printTexto Ini_X + 2.7, PosLinea, "SUCURSAL: " & DireccionEstab
       PosLinea = PosLinea + 0.3
    End If
    cPrint.printTexto Ini_X, PosLinea, "TELEFONOS: " & Telefono1 & "/" & Telefono2 & "/" & FAX
    PosLinea = PosLinea + 0.3
    cPrint.printTexto Ini_X, PosLinea, ULCase(NombreCiudad) & " - Ecuador"
    PosLinea = PosLinea + 0.5
    With TRecibo
         If .Fecha = "" Then .Fecha = FechaSistema
         'PosLinea = 3.4
         cPrint.letraTipo tipoDeLetra, 10
         If .Tipo_Recibo = "I" Then
            cPrint.printTexto Ini_X, PosLinea, "COMPROBANTE DE INGRESO"
         Else
            cPrint.printTexto Ini_X, PosLinea, "COMPROBANTE DE EGRESO"
         End If
         cPrint.printTexto Ini_X + 5.8, PosLinea, "No. " & Format$(Year(.Fecha), "0000") & "-" & Format$(.Recibo_No, "00000000")
         PosLinea = PosLinea + 0.4
         Ini_Y = PosLinea
         cPrint.printCuadro Ini_X - 0.1, Ini_Y, Ini_X + 9, 4.5, Turquesa, "BF"
         PosLinea = PosLinea + 0.1
         cPrint.letraTipo tipoDeLetra, 7

         cPrint.printTexto Ini_X, PosLinea, "Fecha: " & FechaStrg(.Fecha)
         cPrint.printTexto Ini_X + 6, PosLinea, "Por USD"
         cPrint.printVariable Ini_X + 6.5, PosLinea, .Total, True, 3
         PosLinea = PosLinea + 0.5
         cPrint.printTexto Ini_X, PosLinea, "Beneficiario:"
         cPrint.printTexto Ini_X + 2, PosLinea, .Cobrado_a
         PosLinea = PosLinea + 0.5
         cPrint.printTexto Ini_X, PosLinea, "La suma de: " & Cambio_Letras(.Total, 2)
         PosLinea = PosLinea + 0.6
         
         cPrint.printTexto Ini_X, PosLinea, "POR CONCEPTO DE:"
         PosLinea = PosLinea + 0.3
         PosLinea = cPrint.printTextoMultiple(Ini_X, PosLinea, .Concepto, 9)
         cPrint.letraTipo tipoDeLetra, 8
         PosLinea = PosLinea + 0.05
         cPrint.printLinea Ini_X, PosLinea, Ini_X + 8.9, PosLinea, Negro
         PosLinea = PosLinea + 0.1
         cPrint.printTexto Ini_X + 5.5, PosLinea, "TOTAL"
         cPrint.printTexto Ini_X + 6.9, PosLinea, "USD"
         cPrint.printVariable Ini_X + 6.2, PosLinea - 0.05, .Total, True, 3
         PosLinea = PosLinea + 0.35
         cPrint.printTexto Ini_X + 5.5, PosLinea, "ABONADO"
         cPrint.printTexto Ini_X + 6.9, PosLinea, "USD"
         cPrint.printVariable Ini_X + 6.2, PosLinea - 0.05, .SubTotal, True, 3
         PosLinea = PosLinea + 0.35
         cPrint.printTexto Ini_X + 5.5, PosLinea, "SALDO"
         cPrint.printTexto Ini_X + 6.9, PosLinea, "USD"
         cPrint.printVariable Ini_X + 6.2, PosLinea - 0.05, .Saldo, True, 3
         PosLinea = PosLinea + 0.8
         cPrint.printLinea Ini_X + 0.1, PosLinea - 0.1, Ini_X + 2, PosLinea - 0.1, Negro
         cPrint.printLinea Ini_X + 3.5, PosLinea - 0.1, Ini_X + 5, PosLinea - 0.1, Negro
         cPrint.printTexto Ini_X + 0.1, PosLinea, "CONFORME"
         cPrint.printTexto Ini_X + 3.5, PosLinea, "PROCESADO"
         PosLinea = PosLinea + 0.3
         cPrint.printTexto Ini_X + 0.1, PosLinea, "C.I./R.U.C."
         cPrint.printTexto Ini_X + 3.5, PosLinea, "POR"
         PosLinea = PosLinea + 0.3
         cPrint.printTexto Ini_X + 0.1, PosLinea, .CI_RUC
         cPrint.printTexto Ini_X + 3.5, PosLinea, .CodUsuario
         Ini_X = Ini_X - 0.1
         cPrint.printCuadro Ini_X, Ini_Y, Ini_X + 9.1, PosLinea, Negro, "B"
         
         cPrint.printLinea Ini_X, Ini_Y + 0.5, Ini_X + 9.1, Ini_Y + 0.5, Negro
         cPrint.printLinea Ini_X, Ini_Y + 1, Ini_X + 9.1, Ini_Y + 1, Negro
         cPrint.printLinea Ini_X, Ini_Y + 1.5, Ini_X + 9.1, Ini_Y + 1.5, Negro
    End With
    Copia = Copia + 1
    If Copia <= 1 Then GoTo Imprimir_Copia
    cPrint.finalizaImpresion
    RatonNormal
End Sub

Public Sub Imprimir_Recibo_Anticipos(CA As Comprobantes, AbrirDoc As Boolean)
Dim Ini_X As Single
Dim Ini_Y As Single
Dim tipoDeLetra As String
Dim printLinea As String
Dim PrintCar As String
  
    RatonReloj
    tipoDeLetra = TipoCourierNew 'TipoCourier - TipoTimes
   'Geneeramos el documento
    tPrint.TipoImpresion = Es_PDF
    tPrint.NombreArchivo = "Recibo_No_" & CA.TP & "-" & Format$(CA.Numero, "000000000")
    tPrint.TituloArchivo = "Recibo de Abono anticipado"
    tPrint.TipoLetra = tipoDeLetra
    tPrint.OrientacionPagina = 1
    tPrint.PaginaA4 = True
    tPrint.EsCampoCorto = False
    tPrint.VerDocumento = AbrirDoc
    
    Set cPrint = New cImpresion
    cPrint.iniciaImpresion
    Ini_X = 1
    
    cPrint.letraTipo tipoDeLetra, 8
    cPrint.printImagen LogoTipo, Ini_X + 0.1, 1, 3.5, 1.7
    PosLinea = 2.7
    cPrint.colorDeLetra = Negro
    PosLinea = cPrint.printTextoMultiple(Ini_X + 0.1, PosLinea, UCaseStrg(RazonSocial), 6.8)
    PosLinea = PosLinea + 0.35
    If UCaseStrg(RazonSocial) <> UCaseStrg(NombreComercial) Then
       PosLinea = cPrint.printTextoMultiple(Ini_X + 0.1, PosLinea, UCaseStrg(NombreComercial), 6.8)
       PosLinea = PosLinea + 0.35
    End If
    cPrint.printTexto Ini_X + 0.1, PosLinea, "R.U.C.: " & RUC
    PosLinea = PosLinea + 0.35
    cPrint.letraTipo tipoDeLetra, 7
    cPrint.printTexto Ini_X + 0.1, PosLinea, "TELEFONOS: " & Telefono1 & "/" & Telefono2
    PosLinea = PosLinea + 0.35
    cPrint.printTexto Ini_X + 0.1, PosLinea, "DIRECCION: " & Direccion
    PosLinea = PosLinea + 0.35
    If UCaseStrg(Direccion) <> UCaseStrg(DireccionEstab) Then
       cPrint.printTexto Ini_X + 0.1, PosLinea, "SUCURSAL: " & DireccionEstab
       PosLinea = PosLinea + 0.35
    End If
    cPrint.printTexto Ini_X + 0.1, PosLinea, ULCase(NombreCiudad) & " - Ecuador"
    PosLinea = PosLinea + 0.6
    With CA
         cPrint.letraTipo tipoDeLetra, 8
         cPrint.printTexto Ini_X + 0.1, PosLinea, "ABONO ANTICIPADO No."
         cPrint.printTexto Ini_X + 4, PosLinea, Format$(Year(.Fecha), "0000") & "-" & .TP & "-" & Format$(.Numero, "000000000")
         PosLinea = PosLinea + 0.7
         cPrint.printCuadro Ini_X - 0.05, PosLinea, Ini_X + 6.1, PosLinea + 0.5, Magenta, "BF"
         cPrint.printCuadro Ini_X - 0.05, PosLinea, Ini_X + 6.1, PosLinea + 0.5, Negro, "B"
         cPrint.printCuadro Ini_X - 0.05, PosLinea + 0.5, Ini_X + 6.1, PosLinea + 1, Negro, "B"
         cPrint.printCuadro Ini_X - 0.05, PosLinea + 1.05, Ini_X + 6.1, PosLinea + 1.55, Negro, "B"
         cPrint.printCuadro Ini_X - 0.05, PosLinea, Ini_X + 6.1, PosLinea + 3.2, Negro, "B"
         cPrint.printLinea Ini_X - 0.1, PosLinea + 4.5, Ini_X + 3.5, PosLinea + 4.5, Negro
         PosLinea = PosLinea - 0.1
         cPrint.printTexto Ini_X + 0.1, PosLinea, "Fecha: " & FechaStrg(.Fecha)
         cPrint.printTexto Ini_X + 4.4, PosLinea, "Por $"
         cPrint.printVariable Ini_X + 4.3, PosLinea, .Efectivo
         PosLinea = PosLinea + 0.55
         cPrint.letraTipo tipoDeLetra, 7
         cPrint.printTexto Ini_X + 0.1, PosLinea, "Beneficiario:"
         cPrint.printTexto Ini_X + 1.55, PosLinea, .Beneficiario
         PosLinea = PosLinea + 0.5
         cPrint.printTexto Ini_X + 0.1, PosLinea, "La suma de: " & Cambio_Letras(.Efectivo, 2)
         PosLinea = PosLinea + 0.5
         cPrint.printTexto Ini_X + 0.1, PosLinea, "POR CONCEPTO DE:"
         PosLinea = PosLinea + 0.3
         PosLinea = cPrint.printTextoMultiple(Ini_X + 0.1, PosLinea, .Concepto, 7)
         PosLinea = PosLinea + 2.7
         cPrint.printTexto Ini_X + 0.1, PosLinea, "C O N F O R M E"
         PosLinea = PosLinea + 0.35
         cPrint.printTexto Ini_X + 0.1, PosLinea, "C.I./R.U.C. " & .RUC_CI
         PosLinea = PosLinea + 0.35
    End With
    RutaDocumentoPDF = RutaSysBases & "\TEMP\" & tPrint.NombreArchivo & ".pdf"
    cPrint.finalizaImpresion
    RatonNormal
End Sub

Public Sub Imprimir_Apertura(CuentaNo As String)
Dim DatasN As ADODB.Recordset
Dim DatasC As ADODB.Recordset
Dim No_Soc As Byte
On Error GoTo Errorhandler
  CodigoCli = Ninguno
  sSQL = "SELECT * " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Cuenta_No = '" & CuentaNo & "' "
  Select_AdoDB DatasN, sSQL
  With DatasN
   If .RecordCount > 0 Then
       CodigoCli = .fields("Codigo")
       No_Soc = .fields("Num")
       Mifecha = .fields("Fecha_Registro")
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Clientes_Familiares " _
       & "WHERE Codigo = '" & CodigoCli & "' "
Select_AdoDB DatasN, sSQL

Mensajes = "Imprimir Apertura"
Titulo = "Pregunta de Imprimir"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
  RatonReloj
  InicioX = 0.5
  AdoDBAnchoCampos 0.5, DatasN, 10, TipoTimes, 1
   
  sSQL = "SELECT * " _
       & "FROM Tabla_Apertura " _
       & "WHERE ME = 0 "
  Select_AdoDB DatasN, sSQL
  With DatasN
   If .RecordCount > 0 Then
       SaldoDisp = .fields("Monto_Aper")
       SaldoCont = .fields("Monto_Cert")
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Clientes_Familiares " _
       & "WHERE Codigo = '" & CodigoCli & "' "
  Select_AdoDB DatasN, sSQL
  Printer.FontSize = 11
  With DatasN
   If .RecordCount > 0 Then
       PrinterVariables 3, 16, .fields("Nombres")
       PrinterVariables 16.5, 16, .fields("LugarTrabajo")
       PrinterVariables 3, 17.3, .fields("Direccion")
       PrinterVariables 15, 17.3, .fields("Telefono")
       PrinterVariables 16, 18.2, .fields("Cedula")
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Codigo = '" & CodigoCli & "' "
  Select_AdoDB DatasN, sSQL
  With DatasN
       'MsgBox .RecordCount
   If .RecordCount > 0 Then
      'Apertura
       PrinterVariables 3, 3.8, FechaStrg(Mifecha)
       PrinterVariables 12.5, 3.8, Empresa
       PrinterVariables 16, 4.7, CuentaNo
       PrinterVariables 15, 5.9, .fields("Cliente")
       PrinterVariables 1.5, 7, .fields("Representante")
       Cadena = .fields("Cliente") & " " & .fields("Representante")
       PrinterVariables 3, 9.6, Cadena
       PrinterVariables 15, 10.3, .fields("Telefono")
       PrinterVariables 3, 11.1, .fields("Direccion")
       PrinterVariables 3.5, 11.8, .fields("Est_Civil")
       PrinterVariables 15, 11.8, .fields("No_Dep")
       PrinterVariables 3.5, 12.3, .fields("CI_RUC")
       PrinterVariables 15, 12.3, .fields("Ciudad")
       PrinterVariables 4.5, 12.9, .fields("Lugar_Trabajo")
       PrinterVariables 3, 13.5, .fields("DireccionT")
       PrinterVariables 15, 13.5, .fields("TelefonoT")
       PrinterVariables 4.5, 14.2, .fields("Profesion")
      'Razon Social
       PrinterVariables 3.5, 20.8, .fields("Cliente")
       PrinterVariables 14.5, 21.3, .fields("CI_RUC")
       PrinterVariables 3, 21.3, .fields("DireccionT")
       PrinterVariables 3, 21.9, .fields("Ciudad")
       PrinterVariables 10.5, 21.9, .fields("TelefonoT")
       PrinterVariables 14.5, 21.9, .fields("FAX")
       PrinterVariables 4, 22.7, .fields("Casilla")
       Printer.FontSize = 10
      'Autorizacion
       Printer.NewPage
       PrinterVariables 3.5, 3.6, FechaStrg(Mifecha)
       Cadena = .fields("Cliente") & " " _
              & .fields("Representante")
       PrinterVariables 2.5, 4.8, Cadena
       PrinterVariables 2.8, 6.6, CuentaNo
       PrinterVariables 11.2, 6.6, Redondear(SaldoDisp, 2)
       PrinterVariables 9, 7.4, Redondear(SaldoCont * No_Soc, 2) ' Certif
      'Registro Firmas
       Printer.NewPage
       PrinterVariables 2.5, 2.3, .fields("Cliente")
       PrinterVariables 2.8, 3, .fields("Direccion")
       PrinterVariables 2.6, 3.6, .fields("Telefono")
       PrinterVariables 2, 4.4, .fields("CI_RUC")
   End If
  End With
  RatonNormal
  Printer.EndDoc
Else
    RatonNormal
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Function PrinterTextoMaximo(TextoFuente As String, _
                                   LimiteCm As Single) As String
Dim TextoDestino As String
Dim IText As Long
 TextoDestino = Ninguno
 TextoFuente = Replace(TextoFuente, Chr(13) & Chr(10), " ")
 If Len(TextoFuente) > 0 Then
    IText = 1
    TextoDestino = TextoFuente
    If Printer.TextWidth(TextoDestino) > LimiteCm Then
       Do While IText <= Len(TextoFuente)
          TextoDestino = MidStrg(TextoFuente, 1, IText)
          If Printer.TextWidth(TextoDestino) >= LimiteCm Then
             TextoDestino = TextoDestino & "..."
             IText = Len(TextoFuente) + 1
          End If
          IText = IText + 1
       Loop
    End If
 End If
 PrinterTextoMaximo = TextoDestino
End Function

Public Sub PrinterLineas(Xo As Single, _
                         Yo As Single, _
                         Strg As String, _
                         anchoTexto As Single, _
                         Optional AltoLinea As Single)
Dim Inicio As Long
Dim Final As Long
Dim AnchoStrg As Single
Dim PY As Single
Dim PStrg As String
Dim PStrg1 As String
Dim AStrg As String
PY = Yo
If AltoLinea <= 0 Then AltoLinea = Printer.TextHeight(UCaseStrg(MidStrg(Strg, 1, 1))) + 0.05
AltoLinea = Redondear(AltoLinea, 2)
'MsgBox "Antes " & PosLinea
If Yo > 0 Then
   AStrg = TrimStrg(Strg)
   If Len(AStrg) >= 1 Then
      PY = Yo: Inicio = 1: Final = 1: I = 1
      Do While (Len(AStrg) > 1)
         'MsgBox AStrg
         AnchoStrg = Printer.TextWidth(AStrg)
         If AnchoStrg > anchoTexto Then
            AnchoStrg = Printer.TextWidth(MidStrg(AStrg, Inicio, Final))
            While (AnchoStrg <= anchoTexto) And (Final <= Len(AStrg))
                  If Asc(MidStrg(AStrg, Final, 1)) = 10 Then GoTo Salto_Linea
                  Final = Final + 1
                  AnchoStrg = Printer.TextWidth(MidStrg(AStrg, Inicio, Final))
            Wend
Salto_Linea:
            If MidStrg(AStrg, Final, 1) <> " " Then
               While (Final > 1) And (MidStrg(AStrg, Final, 1) <> " ")
                     Final = Final - 1
               Wend
               If Final <= 1 Then Final = Len(AStrg)
            End If
            If Asc(MidStrg(AStrg, Final, 1)) = 10 Then Final = Final + 1
         Else
            Final = Len(AStrg)
         End If
         If Final <= Len(AStrg) Then
            PStrg = TrimStrg(MidStrg(AStrg, Inicio, Final))
            AStrg = TrimStrg(MidStrg(AStrg, Final, Len(AStrg)))
         Else
            PStrg = ""
            AStrg = ""
         End If
         If Asc(MidStrg(AStrg, 1, 1)) = 10 Then AStrg = TrimStrg(MidStrg(AStrg, 2, Len(AStrg)))
         If Len(AStrg) > 0 Then
            If Asc(MidStrg(AStrg, Len(AStrg), 1)) = 10 Then AStrg = TrimStrg(MidStrg(AStrg, 1, Len(AStrg) - 1))
         End If
''         If PosLinea >= LimiteAlto Then
''            'Printer.NewPage
''            PY = 2
''         End If
         'MsgBox PY
         Printer.CurrentX = Xo + 0.1
         Printer.CurrentY = PY
         PStrg1 = " "
         If Len(PStrg) > 0 Then
            For I = 1 To Len(PStrg)
             If Asc(MidStrg(PStrg, I, 1)) <> 10 Then PStrg1 = PStrg1 & MidStrg(PStrg, I, 1)
            Next I
         End If
         PStrg1 = PStrg1 & " "
         PStrg = TrimStrg(PStrg1)
         Printer.Print PStrg
         'MsgBox AltoLinea
         If AltoLinea > 0 Then PY = PY + AltoLinea Else PY = PY + 0.4
         Final = 1
      Loop
   End If
End If
PosLinea = PY + 0.35
' MsgBox "Despues " & PosLinea
End Sub

Public Sub PrinterTextoNegrilla(X1o As Single, _
                                Y1o As Single, _
                                Variable, _
                                Optional Decimales As Byte, _
                                Optional Abreviatura As Boolean)
  If Variable > 0 Then
     If 1 <= Variable And Variable <= 14 Then
        Printer.FontBold = True
        Printer.FontUnderline = True
     End If
     If Abreviatura Then
        Select Case Variable
          Case 1 To 11: PrinterTexto X1o + 0.2, Y1o, "I"
          Case 12 To 13: PrinterTexto X1o + 0.2, Y1o, "R"
          Case 14 To 15: PrinterTexto X1o + 0.2, Y1o, "B"
          Case 16 To 18: PrinterTexto X1o, Y1o, "MB"
          Case 19 To 20: PrinterTexto X1o + 0.2, Y1o, "S"
        End Select
     Else
        If Decimales > 0 Then
           PrinterTexto X1o, Y1o, Format$(Variable, "00." & String$(Decimales, "0"))
        Else
           PrinterTexto X1o, Y1o, Format$(Variable, "00")
        End If
     End If
     If 1 <= Variable And Variable <= 14 Then
        Printer.FontBold = False
        Printer.FontUnderline = False
     End If
  End If
End Sub

Public Function PrinterLineasMayor(Xo As Single, _
                                   Yo As Single, _
                                   Strg As String, _
                                   anchoTexto As Single, _
                                   Optional AltoLinea As Single) As Byte
Dim AnchoStrg As Single
Dim PY As Single
Dim PX As Single
Dim AltoCaracter As Single
Dim PStrg As String
Dim PStrgTemp As String
Dim CStrg As String
Dim NoLineas As Byte
Dim NoLetras As Byte
Dim BlancoLetras As Boolean
Dim Inicio As Long
Dim Final As Long
PY = Yo: NoLineas = 0
If anchoTexto <= 0 Then anchoTexto = 2
If Strg <> Ninguno Then
CStrg = Strg
'MsgBox CStrg
If AltoLinea > 0 Then
   AltoCaracter = AltoLinea
Else
   AltoCaracter = Redondear(Printer.TextHeight("H"), 2)
End If
If Yo > 0 Then
   Do While (Len(CStrg) >= 1)
      'MsgBox Strg
      Inicio = 1: Final = 0: AnchoStrg = 0
      While (AnchoStrg <= anchoTexto) And (Final <= Len(CStrg))
         Final = Final + 1
         AnchoStrg = Printer.TextWidth(MidStrg(CStrg, 1, Final))
      Wend
      PStrgTemp = MidStrg(CStrg, 1, Final)
      BlancoLetras = True
      NoLetras = 0
      While ((BlancoLetras) And (Final > 1))
            Final = Final - 1
            NoLetras = NoLetras + 1
            If MidStrg(CStrg, Final, 1) = " " Then BlancoLetras = False
            If NoLetras >= 20 Then BlancoLetras = False
      Wend
      If Printer.TextWidth(PStrgTemp) < anchoTexto Then
         PStrg = PStrgTemp
         CStrg = ""
      Else
         If Final > 1 Then
            If MidStrg(CStrg, Final, 1) = " " Then
               PStrg = MidStrg(CStrg, 1, Final)
               CStrg = MidStrg(CStrg, Final, Len(CStrg))
            Else
               PStrg = MidStrg(CStrg, 1, Final - 1)
               CStrg = MidStrg(CStrg, Final - 1, Len(CStrg))
            End If
         End If
      End If
      PX = anchoTexto
      'MsgBox TrimStrg(PStrg)
      Printer.Line (Xo, PY)-(Xo + PX, PY + 0.4), Blanco, BF
      'If PY <= LimiteAlto Then
         Printer.CurrentX = Xo + 0.1: Printer.CurrentY = PY
         Printer.Print TrimStrg(PStrg)
      'End If
      PY = PY + AltoCaracter
      NoLineas = NoLineas + 1
      'MsgBox AltoLinea & vbCrLf & TrimStrg(PStrg) & vbCrLf & NoLineas & vbCrLf & PosLinea
   Loop
End If
End If
'PY = PY + 0.35
If NoLineas >= 1 Then PY = PY - AltoCaracter
'MsgBox Yo & " - " & PY
PosLinea_Aux = Redondear(PY, 2)    ' Si no queremos calcular el alto lo coje automaticamente
PrinterLineasMayor = NoLineas
End Function

Public Function PrinterLineasTexto(Xo As Single, _
                                   Yo As Single, _
                                   Strg As String, _
                                   Optional AnchoLinea As Single) As Single
Dim CStrg As String
Dim Yf As Single
Dim Inicio As Long
Dim Final As Long
Dim ChrCR As Boolean
Dim ChrLF As Boolean
Yf = Yo
If Strg <> Ninguno And Yo > 0 And Xo > 0 Then
   RatonReloj
   CStrg = ""
   Inicio = 1: Final = Len(Strg)
   ChrCR = False: ChrLF = False
   Do While Inicio <= Final
      If ChrCR And ChrLF Then
         'MsgBox CStrg
         If Yf < LimiteAlto Then
            Printer.CurrentX = Xo
            Printer.CurrentY = Yf
            Printer.Print CStrg
         End If
         ChrCR = False: ChrLF = False: CStrg = ""
         Yf = Yf + Printer.TextHeight("H") + 0.01
      End If
      If MidStrg(Strg, Inicio, 1) = vbCr Then ChrCR = True
      If MidStrg(Strg, Inicio, 1) = vbLf Then ChrLF = True
      CStrg = CStrg & MidStrg(Strg, Inicio, 1)
      Inicio = Inicio + 1
      If (Printer.TextWidth(CStrg) >= AnchoLinea) And (AnchoLinea > 0) Then
         If MidStrg(Strg, Inicio, 1) = " " Then Inicio = Inicio + 1
         ChrCR = True: ChrLF = True
      End If
   Loop
   If CStrg <> "" Then
      If Yf < LimiteAlto Then
         Printer.CurrentX = Xo
         Printer.CurrentY = Yf
         Printer.Print CStrg
      End If
   End If
   
End If
RatonNormal
PrinterLineasTexto = Yf
End Function

Public Function PrinterLineasTextoPV(Xo As Single, _
                                    Yo As Single, _
                                    Strg As String, _
                                    Optional CantCaracter As Byte) As Single
Dim CStrg As String
Dim cTemp As String
Dim Carac As String
Dim IdCar As Long
Dim IdBlanco As Long
Dim CantBlanco As Long
Dim Yf As Single
Dim Inicio As Long
Dim Final As Long
Yf = Yo
If Strg <> Ninguno And Yo > 0 And Xo > 0 Then
   RatonReloj
   If CantCaracter < 36 Then CantCaracter = 36
   CStrg = Replace(Strg, vbCrLf, "^")
   IdCar = 1
   Inicio = 1: Final = Len(Strg)
   cTemp = ""
   Carac = ""

   Do While Len(CStrg) > 1
      Carac = MidStrg(CStrg, IdCar, 1)
      cTemp = cTemp & Carac
      
      'MsgBox "|" & Carac & "|" & vbCrLf & cTemp
            
      If Len(cTemp) > CantCaracter Or Carac = "^" Then
         Printer.CurrentX = Xo
         Printer.CurrentY = Yf
         
        'MsgBox Asc(Carac) & "Imprimiendo:" & vbCrLf & "|" & cTemp & "|"
         If Carac = "^" Then
            Printer.Print MidStrg(cTemp, 1, Len(cTemp) - 1)
            CStrg = MidStrg(CStrg, Len(cTemp) + 1, Len(CStrg))
         Else
            'MsgBox Asc(Carac) & "Imprimiendo:" & vbCrLf & "|" & cTemp & "|"
''            If Len(cTemp) <> CantCaracter Then
''               CantBlanco = 0
''               IdBlanco = Len(cTemp)
''               Do While IdBlanco > 0 And MidStrg(cTemp, IdBlanco, 1) <> " "
''                  CantBlanco = CantBlanco + 1
''                  IdBlanco = IdBlanco - 1
''               Loop
''               If CantBlanco > 0 Then cTemp = MidStrg(cTemp, 1, IdBlanco)
''            End If
            'MsgBox Asc(Carac) & "Imprimiendo:" & vbCrLf & "|" & cTemp & "|" & vbCrLf & "[" & MidStrg(cTemp, 1, IdBlanco) & "]" & vbCrLf & CantBlanco
            
            Printer.Print TrimStrg(MidStrg(cTemp, 1, Len(cTemp) - 1))
            CStrg = MidStrg(CStrg, Len(cTemp), Len(CStrg))
         End If
         cTemp = ""
         IdCar = 0
         Yf = Yf + Printer.TextHeight("H") + 0.01
      End If
      IdCar = IdCar + 1
   Loop
   If cTemp <> "" Then
      Printer.CurrentX = Xo
      Printer.CurrentY = Yf
      Printer.Print cTemp
      Yf = Yf + Printer.TextHeight("H") + 0.01
   End If
   
End If
RatonNormal
PrinterLineasTextoPV = Yf
End Function

Public Sub CentrarForm(Forms As Form)
Dim PosSup, PosIzq As Single
  'Centrar el formulario
   PosIzq = ((Screen.width - 150 - Forms.width) / 2)
   PosSup = ((Screen.Height - 1950 - Forms.Height) / 2)
   If PosIzq < 0 Then PosIzq = 0
   If PosSup < 0 Then PosSup = 0
   If Forms.BorderStyle = 0 Then PosSup = PosSup - 200
   Forms.Left = PosIzq
   Forms.Top = PosSup
End Sub

Public Sub CentrarFrame(Frames As Frame)
Dim PosSup, PosIzq As Single
  'Centrar el Frame
   PosIzq = ((Screen.width - Frames.width) / 2) - 150
   PosSup = ((Screen.Height - Frames.Height) / 2) - 2000
   If PosIzq < 0 Then PosIzq = 0.1
   If PosSup < 0 Then PosSup = 0.1
   Frames.Left = PosIzq
   Frames.Top = PosSup
End Sub

Public Sub CentrarArribaForm(Forms As Form)
Dim PosSup, PosIzq As Single
   'Centrar el formulario
   PosIzq = ((Screen.width - Forms.width) / 2)
   PosSup = 800
   If PosIzq < 0 Then PosIzq = 0
   If PosSup < 0 Then PosSup = 0
   Forms.Left = PosIzq - 50: Forms.Top = PosSup
End Sub

Public Sub ArribaDerechaForm(Forms As Form)
Dim PosSup, PosIzq As Single
   'Centrar el formulario
   PosIzq = Screen.width - Forms.width - 500
   PosSup = 500
   If PosIzq < 0 Then PosIzq = 0
   If PosSup < 0 Then PosSup = 0
   Forms.Left = PosIzq - 50: Forms.Top = PosSup
End Sub

Public Sub AnchoForm(Forms As Form)
Dim PosSup, PosIzq As Single
   'Centrar el formulario
   Forms.width = Screen.width - 350
''   PosIzq = ((Screen.Width - Forms.Width) / 2) - 20
''   PosSup = ((Screen.Height - Forms.Height) / 2) - 550
   Forms.Left = 0: Forms.Top = 0
End Sub

Public Sub AnchoAltoForm(Forms As Form)
Dim PosSup, PosIzq As Single
   'Centrar el formulario
   Forms.width = Screen.width - 100
   Forms.Height = Screen.Height - 1000
   PosIzq = ((Screen.width - Forms.width) / 2) - 20
   PosSup = ((Screen.Height - Forms.Height) / 2) - 550
   If PosIzq < 0 Then PosIzq = 0
   If PosSup < 0 Then PosSup = 0
   Forms.Left = PosIzq: Forms.Top = PosSup
End Sub

Public Sub Escala_Centimetro(Orientacion As Byte, _
                            TipoLetra As String, _
                            PorteLetra As Integer, _
                            Optional PaginaA4 As Boolean, _
                            Optional EsCampoCorto As Boolean)
'Set Printer = Impresora

 If PorteLetra <= 0 Then PorteLetra = 6
 If SetPapelPRN > 0 Then
    Printer.PaperSize = SetPapelPRN
 Else
    Printer.PaperSize = vbPRPSLetter
 End If
 Printer.ScaleMode = vbCentimeters             'Escala de centimetros
'Orientacion_Pagina normal de la hoja y Limite de impresión a lo largo
 Select Case Orientacion_Pagina
   Case 1: Printer.Orientation = vbPRORPortrait 'Vertical
   Case 2: Printer.Orientation = vbPRORLandscape 'Horizontal
   Case Else: Printer.Orientation = vbPRORPortrait
              Orientacion_Pagina = 1
 End Select
 Printer.DrawWidth = 1                  'Ancho de la línea
 Printer.FontName = TipoLetra           'Tipo de letra en todo el sistema
 Printer.FontSize = PorteLetra          'Porte de la letra default
 Printer.FontBold = False
 Printer.FontItalic = False
 Printer.FontUnderline = False
 LimiteAlto = Redondear(Printer.ScaleHeight - 1, 2)
 LimiteAncho = Redondear(Printer.ScaleWidth - 0.5, 2) 'Limite de impresión a lo ancho
 AnchoPapel = Redondear(Printer.ScaleWidth, 4)        'Ancho de impresion del papel
 SetPapelAncho = Redondear(Printer.ScaleWidth, 4)
 SetPapelLargo = Redondear(Printer.ScaleHeight, 4)
 SetAnchoCampos Printer, EsCampoCorto
End Sub

Public Function ProcesarSeteos(BTP As String) As Boolean
Dim SQLSet As String
Dim ConLineas As Boolean
Dim Items As Long
' Establecemos Espacios y seteos de impresion
    For I = 0 To MaxVect - 1
        SetD(I).PosX = 0
        SetD(I).PosY = 0
        SetD(I).Porte = 9
        SetD(I).Encabezado = "."
    Next I
    ConLineas = False
    SQLSet = "SELECT Lineas " _
           & "FROM Formato " _
           & "WHERE TP = '" & BTP & "' " _
           & "AND Item = '" & NumEmpresa & "' "
    ConectarAdoRecordSet SQLSet
    If AdoReg.RecordCount > 0 Then ConLineas = AdoReg.fields("Lineas")
   'MsgBox "Formato: " & SQLSet
    SQLSet = "SELECT * " _
           & "FROM Seteos_Documentos "
    If ConLineas Then
       SQLSet = SQLSet & "WHERE TP = '" & BTP & "' " _
              & "AND Item = '000' "
    Else
       SQLSet = SQLSet & "WHERE TP = 'P" & BTP & "' " _
              & "AND Item = '" & NumEmpresa & "' "
    End If
    SQLSet = SQLSet & "ORDER BY Campo "
   'MsgBox SQLSet
    ConectarAdoRecordSet SQLSet
    If AdoReg.RecordCount > 0 Then
       AdoReg.MoveLast
       I = AdoReg.RecordCount + 1
       AdoReg.MoveFirst
       Do While Not AdoReg.EOF
          Items = AdoReg.fields("Campo")
          SetD(Items).PosX = Redondear(AdoReg.fields("Pos_X"), 4)
          SetD(Items).PosY = Redondear(AdoReg.fields("Pos_Y"), 4)
          SetD(Items).Porte = Redondear(AdoReg.fields("Porte"), 2)
          SetD(Items).Encabezado = AdoReg.fields("Encabezado")
          If SetD(Items).Encabezado = "" Then SetD(I).Encabezado = Ninguno
          If SetD(Items).Porte <= 0 Then SetD(Items).Porte = 8
          AdoReg.MoveNext
       Loop
    End If
    AdoReg.Close
    ProcesarSeteos = ConLineas
End Function

Public Sub SeteosCtas()
Dim SSQLSeteos As String
Dim AdoCtas As ADODB.Recordset
' Establecemos Espacios y seteos de impresion
  RatonReloj
  Inv_Promedio = False
  PVP_Al_Inicio = False
  Cta_Ret = "0"
  Cta_Ret_IVA = "0"
  Cta_IVA = "0"
  Cta_IVA_Inventario = "0"
  Cta_CxP_Retenciones = "0"
  Cta_Desc = "0"
  Cta_Desc2 = "0"
  Cta_CajaG = "0"
  Cta_General = "0"
  Cta_CajaGE = "0"
  Cta_CajaBA = "0"
  Cta_Gastos = "0"
  Cta_Diferencial = "0"
  Cta_Comision = "0"
  Cta_Mantenimiento = "0"
  Cta_Fondo_Mortuorio = "0"
  Cta_Tarjetas = "0"
  Cta_Del_Banco = "0"
  Cta_Seguro = "0"
  Cta_Seguro_I = "0"
' Consultamos las cuentas de la tabla
  SSQLSeteos = "SELECT * " _
             & "FROM Ctas_Proceso " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY T_No "
  Select_AdoDB AdoCtas, SSQLSeteos
  With AdoCtas
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Cadena = Cadena & .fields("Detalle") & vbCrLf
          Select Case .fields("Detalle")
            Case "Cta_Ret"
                 Cta_Ret = .fields("Codigo")
            Case "Cta_Ret_Egreso"
                 Cta_Ret_Egreso = .fields("Codigo")
            Case "Cta_Ret_IVA"
                 Cta_Ret_IVA = .fields("Codigo")
            Case "Cta_Ret_IVA_Egreso"
                 Cta_Ret_IVA_Egreso = .fields("Codigo")
            Case "Cta_IVA"
                 Cta_IVA = .fields("Codigo")
                 DC_IVA = .fields("DC")
            Case "Cta_Descuentos"
                 Cta_Desc = .fields("Codigo")
                 DC_Desc = .fields("DC")
            Case "Cta_Descuentos_Pronto_Pago"
                 Cta_Desc2 = .fields("Codigo")
            Case "Cta_Caja_General"
                 Cta_General = .fields("Codigo")
                 DC_General = .fields("DC")
            Case "Cta_Caja_GMN"
                 Cta_CajaG = .fields("Codigo")
                 DC_CajaG = .fields("DC")
            Case "Cta_Caja_GME"
                 Cta_CajaGE = .fields("Codigo")
                 DC_CajaGE = .fields("DC")
            Case "Cta_Caja_BAU"
                 Cta_CajaBA = .fields("Codigo")
                 DC_CajaBA = .fields("DC")
            Case "Cta_Gastos"
                 Cta_Gastos = .fields("Codigo")
                 DC_Gastos = .fields("DC")
            Case "Cta_Diferencial_Cambiario"
                 Cta_Diferencial = .fields("Codigo")
                 DC_Diferencial = .fields("DC")
            Case "Cta_SubTotal"
                 Cta_SubTotal = .fields("Codigo")
                 DC_SubTotal = .fields("DC")
            Case "Cta_Comision"
                 Cta_Comision = .fields("Codigo")
                 DC_Comision = .fields("DC")
            Case "Cta_Faltantes"
                 Cta_Faltantes = .fields("Codigo")
                 DC_Faltantes = .fields("DC")
            Case "Cta_Protestos"
                 Cta_Protestos = .fields("Codigo")
                 DC_Protestos = .fields("DC")
            Case "Cta_Sobrantes"
                 Cta_Sobrantes = .fields("Codigo")
                 DC_Sobrantes = .fields("DC")
            Case "Cta_Suspenso"
                 Cta_Suspenso = .fields("Codigo")
                 DC_Suspenso = .fields("DC")
            Case "Cta_Libretas"
                 Cta_Libretas = .fields("Codigo")
                 DC_Libretas = .fields("DC")
            Case "Cta_Certificado"
                 Cta_Certificado = .fields("Codigo")
                 DC_Certificado = .fields("DC")
            Case "Cta_Certificado_Aportacion"
                 Cta_Certificado_Apor = .fields("Codigo")
                 DC_Certificado_Apor = .fields("DC")
            Case "Cta_Apertura"
                 Cta_Apertura = .fields("Codigo")
                 DC_Apertura = .fields("DC")
            Case "Cta_Transito"
                 Cta_Transito = .fields("Codigo")
                 DC_Transito = .fields("DC")
            Case "Cta_Cheque_Transito"
                 Cta_Cheque_Transito = .fields("Codigo")
            Case "Cta_IVA_Inventario"
                 Cta_IVA_Inventario = .fields("Codigo")
                 DC_IVA_Inventario = .fields("DC")
            Case "Cta_CxP_Retenciones"
                 Cta_CxP_Retenciones = .fields("Codigo")
            Case "Cta_Inventario"
                 Cta_Inventario = .fields("Codigo")
                 DC_Inventario = .fields("DC")
            Case "Cta_Mantenimiento"
                 Cta_Mantenimiento = .fields("Codigo")
                 DC_Mantenimiento = .fields("DC")
            Case "Cta_Fondo_Mortuorio"
                 Cta_Fondo_Mortuorio = .fields("Codigo")
                 DC_Fondo_Mortuorio = .fields("DC")
            Case "Cta_Servicios_Basicos"
                 Cta_Servicios_Basicos = .fields("Codigo")
                 DC_Servicios_Basicos = .fields("DC")
            Case "Cta_Servicio"
                 Cta_Servicio = .fields("Codigo")
                 DC_Servicio = .fields("DC")
            Case "Cta_CxP_Propinas"
                 Cta_Propinas = .fields("Codigo")
                 DC_Propinas = .fields("DC")
            Case "Cta_Intereses"
                 Cta_Interes = .fields("Codigo")
                 DC_Interes = .fields("DC")
            Case "Cta_Intereses1"
                 Cta_Interes1 = .fields("Codigo")
                 DC_Interes1 = .fields("DC")
            Case "Cta_CxP_Tarjetas"
                 Cta_Tarjetas = .fields("Codigo")
            Case "Cta_Caja_Vaucher"
                 Cta_Caja_Vaucher = .fields("Codigo")
            Case "Cta_Banco"
                 Cta_Del_Banco = .fields("Codigo")
            Case "Cta_Seguro_Desgravamen"
                 Cta_Seguro = .fields("Codigo")
            Case "Cta_Impuesto_Renta_Empleado"
                 Cta_Impuesto_Renta_Empleado = .fields("Codigo")
            Case "Cta_Seguro_Ingreso"
                 Cta_Seguro_I = .fields("Codigo")
            Case "Inv_Promedio": If .fields("Codigo") = "TRUE" Then Inv_Promedio = True
            Case "PVP_Al_Inicio": If .fields("Codigo") = "TRUE" Then PVP_Al_Inicio = True
          End Select
         .MoveNext
       Loop
   End If
  End With
  AdoCtas.Close
  SSQLSeteos = ""
 'If Cta_Ret = "0" Then SSQLSeteos = SSQLSeteos & "Cta_Ret_Ingreso" & vbCrLf
  If Cta_IVA = "0" Then SSQLSeteos = SSQLSeteos & "Cta_IVA" & vbCrLf
  If Cta_Desc = "0" Then SSQLSeteos = SSQLSeteos & "Cta_Descuentos" & vbCrLf
  If Cta_Desc2 = "0" Then SSQLSeteos = SSQLSeteos & "Cta_Descuentos_Pronto_Pago" & vbCrLf
  If Cta_CajaG = "0" Then SSQLSeteos = SSQLSeteos & "Cta_Caja_GMN" & vbCrLf
  If Cta_General = "0" Then SSQLSeteos = SSQLSeteos & "Cta_Caja_General" & vbCrLf
  If Cta_CajaGE = "0" Then SSQLSeteos = SSQLSeteos & "Cta_Caja_GME" & vbCrLf
  If Cta_CajaBA = "0" Then SSQLSeteos = SSQLSeteos & "Cta_Caja_VAU" & vbCrLf
  If Cta_Gastos = "0" Then SSQLSeteos = SSQLSeteos & "Cta_Gastos" & vbCrLf
  If Cta_Diferencial = "0" Then SSQLSeteos = SSQLSeteos & "Cta_Diferencial_Cambiario" & vbCrLf
  If Cta_IVA_Inventario = "0" Then SSQLSeteos = SSQLSeteos & "Cta_IVA_Inventario" & vbCrLf
  RatonNormal
  If SSQLSeteos <> "" Then
     SSQLSeteos = "Verifique el codigo de:" & vbCrLf & SSQLSeteos & vbCrLf _
                & "La proxima vez que ejecute el sistema se crearan estas cuentas "
   ' MsgBox SSQLSeteos
   ' CtasSeteos AdoRec
  End If
  
End Sub

Public Function Leer_Seteos_Ctas(Det_Cta As String) As String
Dim SSQLSeteos As String
Dim Cta_Ret_Aux As String
Dim AdoCtas As ADODB.Recordset
    RatonReloj
    Cta_Ret_Aux = "0"
    SSQLSeteos = "SELECT * " _
               & "FROM Ctas_Proceso " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Detalle = '" & Det_Cta & "' "
    Select_AdoDB AdoCtas, SSQLSeteos
    If AdoCtas.RecordCount > 0 Then Cta_Ret_Aux = AdoCtas.fields("Codigo")
    AdoCtas.Close
    Leer_Seteos_Ctas = Cta_Ret_Aux
    RatonNormal
End Function

'''Public Sub SeteosCodigos()
'''Dim LstSN As Boolean
'''Dim SQLAux As String
'''Dim AdoDB000 As ADODB.Recordset
'''Dim AdoDBItem As ADODB.Recordset
'''
'''  SQLAux = "SELECT * " _
'''         & "FROM Codigos " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "ORDER BY Concepto "
'''  Select_AdoDB AdoDBItem, SQLAux
'''
'''  SQLAux = "SELECT * " _
'''         & "FROM Codigos " _
'''         & "WHERE Item = '000' " _
'''         & "AND Periodo = '.' " _
'''         & "ORDER BY Concepto "
'''  Select_AdoDB AdoDB000, SQLAux
'''  RatonReloj
'''  With AdoDB000
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Codigo = .Fields("Concepto")
'''          If AdoDBItem.RecordCount > 0 Then
'''             AdoDBItem.MoveFirst
'''             AdoDBItem.Find ("Concepto = '" & Codigo & "'")
'''             If AdoDBItem.EOF Then
'''                SetAdoAddNew "Codigos"
'''                SetAdoFields "Concepto", Codigo
'''                SetAdoFields "Numero", 1
'''                SetAdoFields "Item", NumEmpresa
'''                SetAdoFields "Periodo", Periodo_Contable
'''                SetAdoUpdate
'''             End If
'''          Else
'''             SetAdoAddNew "Codigos"
'''             SetAdoFields "Concepto", Codigo
'''             SetAdoFields "Numero", 1
'''             SetAdoFields "Item", NumEmpresa
'''             SetAdoFields "Periodo", Periodo_Contable
'''             SetAdoUpdate
'''          End If
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''  AdoDB000.Close
'''  AdoDBItem.Close
'''
'''  SQLAux = "SELECT * " _
'''         & "FROM Ctas_Proceso " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "ORDER BY T_No "
'''  Select_AdoDB AdoDBItem, SQLAux
'''
'''  SQLAux = "SELECT * " _
'''         & "FROM Ctas_Proceso " _
'''         & "WHERE Item = '000' " _
'''         & "AND Periodo = '.' " _
'''         & "ORDER BY T_No "
'''  Select_AdoDB AdoDB000, SQLAux
'''  RatonReloj
'''  With AdoDB000
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          TipoDoc = .Fields("DC")
'''          Codigo1 = .Fields("Detalle")
'''          Codigo = .Fields("Codigo")
'''          Numero = .Fields("T_No")
'''          If AdoDBItem.RecordCount > 0 Then
'''             AdoDBItem.MoveFirst
'''             AdoDBItem.Find ("Detalle = '" & Codigo1 & "'")
'''             If AdoDBItem.EOF Then
'''                SetAdoAddNew "Ctas_Proceso"
'''                SetAdoFields "Detalle", Codigo1
'''                SetAdoFields "Codigo", Codigo
'''                SetAdoFields "T_No", Numero
'''                SetAdoFields "DC", TipoDoc
'''                SetAdoFields "Item", NumEmpresa
'''                SetAdoFields "Periodo", Periodo_Contable
'''                SetAdoUpdate
'''             End If
'''          Else
'''             SetAdoAddNew "Ctas_Proceso"
'''             SetAdoFields "Detalle", Codigo1
'''             SetAdoFields "Codigo", Codigo
'''             SetAdoFields "DC", TipoDoc
'''             SetAdoFields "T_No", Numero
'''             SetAdoFields "Item", NumEmpresa
'''             SetAdoFields "Periodo", Periodo_Contable
'''             SetAdoUpdate
'''          End If
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
'''  AdoDB000.Close
'''  AdoDBItem.Close
'''
'''  SQLAux = "SELECT * " _
'''         & "FROM Catalogo_SubCtas " _
'''         & "WHERE Item = '" & NumEmpresa & "' " _
'''         & "AND Periodo = '" & Periodo_Contable & "' " _
'''         & "AND Codigo = '.' " _
'''         & "AND TC = 'GC' "
'''  Select_AdoDB AdoDBItem, SQLAux
'''  With AdoDBItem
'''   If .RecordCount <= 0 Then
'''       SetAdoAddNew "Catalogo_SubCtas"
'''       SetAdoFields "Codigo", Ninguno
'''       SetAdoFields "Detalle", Ninguno
'''       SetAdoFields "TC", "GC"
'''       SetAdoFields "Caja", True
'''       SetAdoFields "Item", NumEmpresa
'''       SetAdoUpdate
'''   End If
'''  End With
'''  AdoDBItem.Close
'''  RatonNormal
'''End Sub

Public Function BlancoCodigoCta() As String
Dim Strg As String
  Strg = ""
  For I = 1 To Len(FormatoCtas)
    If MidStrg(FormatoCtas, I, 1) = "#" Then Strg = Strg & " " Else Strg = Strg & "."
  Next I
  BlancoCodigoCta = Strg
End Function

Public Function CambioCodigoCta(Codigo As String) As String
Dim LongCta As Long
Dim Bandera As Boolean
Dim Codigo_Cta As String
  Bandera = True: LongCta = Len(Codigo)
  Do While ((LongCta > 0) And Bandera)
     If ((MidStrg(Codigo, LongCta, 1) <> ".") And (MidStrg(Codigo, LongCta, 1) <> " ")) Then Bandera = False
     LongCta = LongCta - 1
  Loop
  LongCta = LongCta + 1
  If LongCta < 1 Then LongCta = 1
  Codigo_Cta = MidStrg(Codigo, 1, LongCta)
  If Codigo_Cta = "" Then Codigo_Cta = "0"
  If Codigo_Cta = " " Then Codigo_Cta = "0"
  If Codigo_Cta = Ninguno Then Codigo_Cta = "0"
  CambioCodigoCta = Codigo_Cta
End Function

Public Function CambioCodigoCtaSup(Codigo As String) As String
Dim LongCta As Long
Dim Bandera As Boolean
Dim CadAux As String
  CadAux = "": Bandera = True
  LongCta = Len(Codigo)
  Do While ((LongCta > 0) And Bandera)
     If MidStrg(Codigo, LongCta, 1) = "." Then Bandera = False
     LongCta = LongCta - 1
  Loop
  If LongCta < 1 Then CadAux = "0" Else CadAux = MidStrg(Codigo, 1, LongCta)
  CambioCodigoCtaSup = CadAux
End Function

Public Function CodigoCuentaSup(CodigoCta As String) As String
Dim LongCta As Long
Dim Bandera As Boolean
Dim CadAux As String
  CadAux = CodigoCta: Bandera = True
  LongCta = Len(CadAux)
  Do While ((LongCta > 0) And Bandera)
     If MidStrg(CadAux, LongCta, 1) = "." Then Bandera = False
     LongCta = LongCta - 1
  Loop
  If LongCta < 1 Then CadAux = "0" Else CadAux = MidStrg(CadAux, 1, LongCta)
  CodigoCuentaSup = CadAux
End Function

Public Function FormatoCodigoCta(Cta As String) As String
Dim Ctas, Strg, ch As String
  Ctas = Cta
  If Ctas = Ninguno Then Ctas = "0"
  Strg = "": Cadena = ""
  If Len(Ctas) < 20 Then
     For I = 1 To Len(MascaraCtas)
        ch = MidStrg(MascaraCtas, I, 1)
        If ch = "#" Then Strg = Strg & " " Else Strg = Strg & "."
     Next I
     Cadena = Ctas & MidStrg(Strg, Len(Ctas) + 1, Len(Strg) - Len(Ctas))
  Else
     Cadena = Ctas
  End If
  FormatoCodigoCta = Cadena
End Function

Public Function FormatoCodigoCurso(Cta As String) As String
Dim Ctas, Strg, ch As String
  Ctas = Cta
  If Ctas = Ninguno Then Ctas = "0"
  Strg = "": Cadena = ""
  If Len(Ctas) <= 10 Then
     For I = 1 To Len(MascaraCurso)
        ch = MidStrg(MascaraCurso, I, 1)
        If ch = "C" Then Strg = Strg & " " Else Strg = Strg & "."
     Next I
     Cadena = Ctas & MidStrg(Strg, Len(Ctas) + 1, Len(Strg) - Len(Ctas))
  Else
     Cadena = Ctas
  End If
  FormatoCodigoCurso = Cadena
End Function

Public Function FormatoCodigoActivo(Cta As String) As String
Dim Ctas, Strg, ch As String
  Ctas = Cta
  If Ctas = Ninguno Then Ctas = "0"
  Strg = "": Cadena = ""
  If Len(Ctas) < 20 Then
     For I = 1 To Len(MascaraCodigoA)
        ch = MidStrg(MascaraCodigoA, I, 1)
        If ch = "C" Then Strg = Strg & " " Else Strg = Strg & "."
     Next I
     Cadena = Ctas & MidStrg(Strg, Len(Ctas) + 1, Len(Strg) - Len(Ctas))
  Else
     Cadena = Ctas
  End If
 'MsgBox MascaraCodigoA & vbCrLf & Cadena
  FormatoCodigoActivo = Cadena
End Function

Public Function FormatoCodigoCta6(Cta As String) As String
Dim Ctas, Strg, ch As String
  Ctas = Cta
  If Ctas = Ninguno Then Ctas = "0"
  Strg = "": Cadena = ""
  If Len(Ctas) < 7 Then
     For I = 1 To Len(MascaraCtas6)
        ch = MidStrg(MascaraCtas6, I, 1)
        If ch = "#" Then Strg = Strg & " " Else Strg = Strg & "."
     Next I
     Cadena = Ctas & MidStrg(Strg, Len(Ctas) + 1, Len(Strg) - Len(Ctas))
  Else
     Cadena = Ctas
  End If
  FormatoCodigoCta6 = Cadena
End Function

Public Function FormatoCodigoCta5(Cta As String) As String
Dim Ctas, Strg, ch As String
  Ctas = Cta
  If Ctas = Ninguno Then Ctas = "0"
  Strg = "": Cadena = ""
  If Len(Ctas) < 7 Then
     For I = 1 To Len(MascaraCtas5)
        ch = MidStrg(MascaraCtas5, I, 1)
        If ch = "#" Then Strg = Strg & " " Else Strg = Strg & "."
     Next I
     Cadena = Ctas & MidStrg(Strg, Len(Ctas) + 1, Len(Strg) - Len(Ctas))
  Else
     Cadena = Ctas
  End If
  FormatoCodigoCta5 = Cadena
End Function

Public Function FormatoCodigoRUC_CI(FCta As String, BRUC As Boolean) As String
Dim TCtas, Strg, ch As String
 '123456789-1-123
  TCtas = MidStrg(FCta, 1, 15)
  Strg = "": Cadena = ""
  If BRUC Then
     For I = 1 To Len(MascaraRUC)
         ch = MidStrg(MascaraRUC, I, 1)
         If ch = "#" Then
            If IsNumeric(MidStrg(TCtas, I, 1)) Then
               Strg = Strg & MidStrg(TCtas, I, 1)
            Else
               Strg = Strg & "0"
            End If
         Else
            Strg = Strg & "-"
         End If
     Next I
     Cadena = TrimStrg(MidStrg(Strg, 1, Len(MascaraRUC)))
  Else
     For I = 1 To Len(MascaraCI)
         ch = MidStrg(MascaraCI, I, 1)
         If ch = "#" Then
            If IsNumeric(MidStrg(TCtas, I, 1)) Then
               Strg = Strg & MidStrg(TCtas, I, 1)
            Else
               Strg = Strg & "0"
            End If
         Else
            Strg = Strg & "-"
         End If
     Next I
     Cadena = TrimStrg(MidStrg(Strg, 1, Len(MascaraCI)))
  End If
  FormatoCodigoRUC_CI = Cadena
End Function

Public Function FormatoCodigoTelef(FCta As String) As String
Dim TCtas, Strg, ch As String
 '123456789-1-123
  TCtas = MidStrg(FCta, 1, 10)
  Strg = "": Cadena = ""
  For I = 1 To Len(MascaraTelefC)
      ch = MidStrg(MascaraTelefC, I, 1)
      If ch = "#" Then
         If IsNumeric(MidStrg(TCtas, I, 1)) Then
            Strg = Strg & MidStrg(TCtas, I, 1)
         Else
            Strg = Strg & "0"
         End If
      Else
         Strg = Strg & "-"
      End If
  Next I
  Cadena = TrimStrg(MidStrg(Strg, 1, Len(MascaraTelefC)))
  FormatoCodigoTelef = Cadena
End Function

Public Sub FormatoMaskCta6(FormatMaskCta As MaskEdBox)
   FormatMaskCta.PromptChar = " "
   FormatMaskCta.mask = MascaraCtas6
   FormatMaskCta.Format = FormatoCtas6
End Sub

Public Sub FormatoMaskCta5(FormatMaskCta As MaskEdBox)
   FormatMaskCta.PromptChar = " "
   FormatMaskCta.mask = MascaraCtas5
   FormatMaskCta.Format = FormatoCtas5
End Sub

Function CodigoMaskCta(CodigoCta As String) As String
Dim Cad As String
Cad = ""
If Len(CodigoCta) > 0 Then
   If Len(CodigoCta) < Len(FormatoCtas) Then
      For I = Len(CodigoCta) + 1 To Len(FormatoCtas)
          If MidStrg(FormatoCtas, I, 1) = "." Then Cad = Cad & "."
          If MidStrg(FormatoCtas, I, 1) = "#" Then Cad = Cad & " "
      Next I
   End If
   CodigoMaskCta = CodigoCta & Cad
Else
   For I = 1 To Len(FormatoCtas)
       If MidStrg(FormatoCtas, I, 1) = "." Then Cad = Cad & "."
       If MidStrg(FormatoCtas, I, 1) = "#" Then Cad = Cad & " "
   Next I
   CodigoMaskCta = Cad
End If
End Function

Public Function CambioCodigoKardex(Codigo As String) As String
Dim LongCta As Long
Dim Bandera As Boolean
  Bandera = True: LongCta = Len(Codigo)
  Do While ((LongCta > 0) And Bandera)
     If ((MidStrg(Codigo, LongCta, 1) <> ".") And (MidStrg(Codigo, LongCta, 1) <> " ")) Then Bandera = False
     LongCta = LongCta - 1
  Loop
  LongCta = LongCta + 1
  If LongCta < 1 Then LongCta = 1
  CambioCodigoKardex = MidStrg(Codigo, 1, LongCta)
End Function

Public Function FormatoCodigoCredito(Cta As String) As String
Dim Ctas, Strg, ch As String
  Ctas = Cta
  Strg = "": Cadena = ""
  If Len(Ctas) <= Len(MascaraCodigoC) Then
     For I = 1 To Len(MascaraCodigoC)
        ch = MidStrg(MascaraCodigoC, I, 1)
        If ch = "C" Then Strg = Strg & " " Else Strg = Strg & "."
     Next I
     Cadena = Ctas & MidStrg(Strg, Len(Ctas) + 1, Len(Strg) - Len(Ctas))
  Else
     Cadena = Ctas
  End If
  'MsgBox Cadena
  FormatoCodigoCredito = Cadena
End Function

Public Function FormatoCodigoKardex(Cta As String) As String
Dim Ctas, Strg, ch As String
  Ctas = Cta
  Strg = "": Cadena = ""
  If Len(Ctas) <= Len(MascaraCodigoK) Then
     For I = 1 To Len(MascaraCodigoK)
        ch = MidStrg(MascaraCodigoK, I, 1)
        If ch = "C" Then Strg = Strg & " " Else Strg = Strg & "."
     Next I
     Cadena = Ctas & MidStrg(Strg, Len(Ctas) + 1, Len(Strg) - Len(Ctas))
  Else
     Cadena = Ctas
  End If
  FormatoCodigoKardex = Cadena
End Function

Public Sub FormatoMaskCta(FormatMaskCta As MaskEdBox)
   FormatMaskCta.PromptChar = " "
   FormatMaskCta.mask = MascaraCtas
   FormatMaskCta.Format = FormatoCtas
End Sub

Public Sub FormatoMaskCodK(FormatMaskCta As MaskEdBox)
   FormatMaskCta.PromptChar = " "
   FormatMaskCta.mask = MascaraCodigoK
   FormatMaskCta.Format = FormatoCodigoK
End Sub

Public Sub FormatoMaskCodC(FormatMaskCta As MaskEdBox)
   FormatMaskCta.PromptChar = " "
   FormatMaskCta.mask = MascaraCodigoC
   FormatMaskCta.Format = FormatoCodigoC
End Sub

Public Sub FormatoMaskCodA(FormatMaskCta As MaskEdBox)
   'MsgBox MascaraCodigoA
   FormatMaskCta.PromptChar = " "
   FormatMaskCta.mask = MascaraCodigoA
   FormatMaskCta.Format = FormatoCodigoA
End Sub

Public Sub FormatoMaskCurso(FormatMaskCta As MaskEdBox)
   FormatMaskCta.PromptChar = " "
   FormatMaskCta.mask = MascaraCurso
   FormatMaskCta.Format = FormatoCurso
End Sub

Function CodigoMaskKardex(CodigoCta As String) As String
Dim Cad As String
Cad = ""
If Len(CodigoCta) > 0 Then
   If Len(CodigoCta) < Len(MascaraCodigoK) Then
      For I = Len(CodigoCta) + 1 To Len(MascaraCodigoK)
          If MidStrg(MascaraCodigoK, I, 1) = "." Then Cad = Cad & "."
          If MidStrg(MascaraCodigoK, I, 1) = "#" Then Cad = Cad & " "
      Next I
   End If
   CodigoMaskKardex = CodigoCta & Cad
Else
   For I = 1 To Len(FormatoCtas)
       If MidStrg(MascaraCodigoK, I, 1) = "." Then Cad = Cad & "."
       If MidStrg(MascaraCodigoK, I, 1) = "#" Then Cad = Cad & " "
   Next I
   CodigoMaskKardex = Cad
End If
End Function

Public Sub RatonReloj()
  Screen.MousePointer = vbHourglass
End Sub

Public Sub RatonNormal()
  Screen.MousePointer = vbDefault
End Sub

Public Sub MarcarTexto(Objeto As Control)
  Objeto.SelStart = 0
  Objeto.SelLength = Len(Objeto.Text)
End Sub

Public Function Copiar_Archivos(origen$, destino$, Archivo$)
' Copia varios archivos de una carpeta a otra
' Origen$= directorio de origen , terminado en "\"
' Destino$= directorio de destino , terminado en "\"
' archivo$= especificacion de archivos a copiar, con simb. comodin
' informa= un label en el que se muestra el progreso
'
' result = xfilecopy("c:\pat\", "h:\doc\", "*.exe", label1)
' copia todos los archivos exe de c:\pat en h:\doc
' muestra lo que esta haciendo en label1
Dim N, Result, Cuenta, pcent
' cuenta los archivos a copiar
If MidStrg(origen$, Len(origen$), 1) <> "\" Then origen$ = origen$ & "\"
If MidStrg(destino$, Len(destino$), 1) <> "\" Then destino$ = destino$ & "\"
If Archivo$ = "" Then Archivo$ = "*.*"
RatonReloj

Cuenta = 0
N = Dir$(origen$ & Archivo$)
While (N <> "")
 Cuenta = Cuenta + 1
 N = Dir$
Wend
Progreso_Barra.Mensaje_Box = "Empezando a copiar archivos"
Progreso_Iniciar
Progreso_Barra.Valor_Maximo = Cuenta
' Copia
Result = 0
pcent = 0
N = Dir$(origen$ & Archivo$)
On Error GoTo malxfilecopy
While (N <> "") And (Result > -1)
    pcent = pcent + 1
    Progreso_Barra.Mensaje_Box = "Copiando " & origen$ & N & " a " & destino$
    Progreso_Esperar
    DoEvents

    FileCopy origen$ & N, destino$ & N
    Result = Result + 1
    N = Dir$
continuaxfilecopy:
Wend
'informa.Caption = ""
Copiar_Archivos = Result
RatonNormal
Progreso_Final
Exit Function

malxfilecopy:
 RatonNormal
 Progreso_Final
 Result = -1
 Resume continuaxfilecopy
End Function


Public Function CambiarMarcarTexto(Objeto As Control) As String
Dim CadAux1 As String
Dim CadAux2 As String
  CadAux1 = MidStrg(Objeto.Text, Objeto.SelStart + 1, Objeto.SelLength)
  CadAux2 = MidStrg(Objeto.Text, 1, Objeto.SelStart)
  CambiarMarcarTexto = CadAux1 & " " & CadAux2
End Function

Public Sub MarcarTextoFinal(Objeto As Control)
  Objeto.SelStart = Len(Objeto.Text)
End Sub

Public Function SumaHora(CHora As String, Horas As Single) As String
Dim THor, TMin, TSeg As Single
Dim IHor, IMin, ISeg As Single
  If CHora = "" Then CHora = HoraSistema
  CHora = Format$(CHora, FormatoTimes)
  IHor = Int(Horas)
  IMin = CInt(MidStrg(Format$((Horas - Int(Horas)) * 10000, "0000"), 1, 2))
  ISeg = CInt(MidStrg(Format$((Horas - Int(Horas)) * 10000, "0000"), 3, 2))
  THor = DatePart("h", CHora)
  TMin = DatePart("n", CHora)
  TSeg = DatePart("s", CHora)
  If (TSeg + ISeg) >= 60 Then
     TSeg = (TSeg + ISeg) - 60
     TMin = TMin + 1
  Else
     TSeg = TSeg + ISeg
  End If
  If (TMin + IMin) >= 60 Then
     TMin = (TMin + IMin) - 60
     THor = THor + 1
  Else
     TMin = TMin + IMin
  End If
  THor = THor + IHor
  If THor >= 24 Then THor = THor - 24
  SumaHora = Format$(THor, "00") & ":" & Format$(TMin, "00") & ":" & Format$(TSeg, "00")
End Function

Public Function CHoraLong(CHora As String) As Long
  If CHora = "" Then CHora = HoraSistema
  CHora = Format$(CHora, FormatoTimes)
  CHoraLong = CLng(Format$(DatePart("h", CHora), "00") & Format$(DatePart("n", CHora), "00") & Format$(DatePart("s", CHora), "00"))
End Function

Public Function CLongHora(CHora As Long) As String
Dim CadStr As String
  CadStr = Format$(CHora, "000000")
  CLongHora = MidStrg(CadStr, 1, 2) & ":" & MidStrg(CadStr, 3, 2) & ":" & MidStrg(CadStr, 5, 2)
End Function

Public Function CFechaLong(CFecha As String) As Long
  If Len(CFecha) = 10 Then
     If CFecha = "00/00/0000" Then CFecha = FechaSistema
  Else
     CFecha = FechaSistema
  End If
  If Not IsDate(CFecha) Then CFecha = FechaSistema
  CFecha = Format$(CFecha, FormatoFechas)
  CFechaLong = CLng(CDate(CFecha))
End Function

Public Function CLongFecha(CFecha As Long) As String
  CLongFecha = Format$(CDate(CFecha), FormatoFechas)
End Function

Public Sub PresionoEnter(KeyCode)
  If KeyCode = vbKeyReturn Then Pulsar_Tecla (vbKeyTab)    ' SendKeys "{TAB}", False
End Sub

Public Sub TextoValido(TextB As TextBox, _
                       Optional Numero As Boolean, _
                       Optional Mayusculas As Boolean, _
                       Optional NumeroDecimales As Byte)
Dim TextosB As String
    TextosB = TextB
    If IsNull(TextosB) Then TextosB = ""
    If IsEmpty(TextosB) Then TextosB = ""
    TextosB = Replace(TextosB, vbCr, "")
    TextosB = Replace(TextosB, vbLf, "")
    TextosB = TrimStrg(TextosB)
    If Mayusculas Then TextosB = UCaseStrg(TextosB)
    If Numero Then
       If TextosB = "" Then TextosB = "0"
      'MsgBox IsNumeric(TextosB)
       If IsNumeric(TextosB) Then
          Select Case NumeroDecimales
            Case 0: TextosB = Format$(TextosB, "##0.00")
            Case Is > 2: TextosB = Format$(TextosB, "#,##0." & String$(NumeroDecimales, "0"))
            Case Else: TextosB = Format$(TextosB, "##0.00")
          End Select
          TextB = TrimStrg(TextosB)
       Else
          TextosB = "0"
          TextB = TextosB
          TextB.SetFocus
       End If
    Else
       If TextosB = "" Then TextosB = Ninguno
       TextB = TextosB
    End If
End Sub

Public Function Validar_Texto(TextoValidar As Variant) As String
   If IsNull(TextoValidar) Or IsEmpty(TextoValidar) Then Validar_Texto = "" Else Validar_Texto = TrimStrg(TextoValidar)
End Function

Public Sub ControlEsNumerico(elControl As TextBox)
  elControl.RightToLeft = True
End Sub

Public Sub EncabezadoEmpresa(InicX As Single)
Dim Son_Iguales As Boolean
   Son_Iguales = False
   PorteLetra = Printer.FontSize
   LetraAnterior = Printer.FontName
   Printer.FontName = TipoArialNarrow
   Printer.FontBold = True
   If UCaseStrg(Empresa) = UCaseStrg(NombreComercial) Then Son_Iguales = True
   PosLinea = InicX - 0.1
   Printer.FontSize = 13
   If Son_Iguales Then
      PrinterTexto CentrarTexto(Empresa) - 0.5, PosLinea, Empresa
   Else
      PrinterTexto CentrarTexto(Empresa) - 0.5, PosLinea, Empresa
      If Len(NombreComercial) > 1 Then
         Printer.FontSize = 12
         PosLinea = PosLinea + 0.47
         PrinterTexto CentrarTexto(NombreComercial) - 0.5, PosLinea, NombreComercial
      End If
   End If
   PosLinea = PosLinea + 0.46
   Printer.FontName = TipoArial
   Printer.FontSize = 9
   Cadena = "R.U.C. " & RUC
   PrinterTexto CentrarTexto(Cadena) - 0.5, PosLinea, Cadena
   Printer.FontSize = 7
   Printer.FontName = TipoArialNarrow
   Cadena = "Dir.: " & ULCase(Direccion)
   If Len(Telefono1) > 1 And Val(Telefono1) > 0 Then Cadena = Cadena & " - Teléf.: " & Telefono1
   If Len(FAX) > 1 And Val(FAX) > 0 Then Cadena = Cadena & "/FAX: " & FAX
   PosLinea = PosLinea + 0.35
   PrinterTexto CentrarTexto(Cadena) - 0.5, PosLinea, Cadena
   Printer.FontBold = False
   Printer.FontSize = PorteLetra
   Printer.FontName = TipoArial
End Sub

Public Sub Encabezado_Documento(InicX As Single, _
                                InicY As Single, _
                                AnchoDeLinea As Single, _
                                Optional Titulo_Documento As String)
Dim PosXX As Single
   PorteLetra = Printer.FontSize
   LetraAnterior = Printer.FontName
   Printer.FontName = TipoTimes
   Printer.FontBold = True
   PosLinea = InicY
   PrinterPaint LogoTipo, InicX, PosLinea, 4, 2
   PrinterPaint LogoTipo1, AnchoDeLinea - 4, PosLinea, 4, 2
   If Titulo_Documento <> "" Then
      If Len(Institucion1 & Institucion2) > 2 Then
         Printer.FontSize = 14
         Cadena = UCaseStrg(Institucion1)
         PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
         PosLinea = PosLinea + 0.6
         Printer.FontSize = 12
         Cadena = UCaseStrg(Institucion2)
         PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
         PosLinea = PosLinea + 0.6
      Else
         Printer.FontSize = 22
         Cadena = UCaseStrg(Empresa)
         PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
         PosLinea = PosLinea + 0.9
      End If
'''      Printer.FontSize = 10
'''      Cadena = Direccion
'''      PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
   Else
      Printer.FontSize = 12
      Cadena = "REPÚBLICA DEL ECUADOR"
      PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
      PosLinea = PosLinea + 0.5
      Cadena = "MINISTERIO DE EDUCACIÓN"
      PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
      PosLinea = PosLinea + 0.5
      If UCaseStrg(Empresa) <> UCaseStrg(NombreComercial) Then
         Printer.FontSize = 12
         Printer.ForeColor = Gris
         PrinterTexto CentrarTexto(UCaseStrg(Empresa)) + 0.025, PosLinea, UCaseStrg(Empresa)
         PosLinea = PosLinea + 0.025
         Printer.ForeColor = Negro
         PrinterTexto CentrarTexto(UCaseStrg(Empresa)) + 0.025, PosLinea, UCaseStrg(Empresa)
         PosLinea = PosLinea + 0.5
         PrinterTexto CentrarTexto(NombreComercial), PosLinea, NombreComercial
         PosLinea = PosLinea + 0.5
         Printer.FontSize = 8: Printer.FontItalic = False
         Cadena = Direccion & ". Teléfono: " & Telefono1 & ". FAX: " & FAX
         PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
         PosLinea = PosLinea + 0.5
      Else
         Printer.FontSize = 12
         Printer.ForeColor = Gris
         PrinterTexto CentrarTexto(UCaseStrg(Empresa)) + 0.025, PosLinea, UCaseStrg(Empresa)
         PosLinea = PosLinea + 0.025
         Printer.ForeColor = Negro
         PrinterTexto CentrarTexto(UCaseStrg(Empresa)), PosLinea, UCaseStrg(Empresa)
         PosLinea = PosLinea + 0.5
         Printer.FontBold = False
         Printer.FontSize = 8: Printer.FontItalic = False
         Cadena = Direccion & ". Teléfono: " & Telefono1 & ". FAX: " & FAX
         PrinterTexto CentrarTexto(Cadena), PosLinea, Cadena
         PosLinea = PosLinea + 0.4
      End If
   End If
   Printer.FontBold = True
   If SQLMsg1 <> "" Then
      Printer.FontSize = 10
      PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
      PosLinea = PosLinea + 0.4
   End If
   Printer.FontBold = False
   If SQLMsg2 <> "" Then
      Printer.FontSize = 9
      PosXX = LimiteAncho - Printer.TextWidth(SQLMsg2) - 2
      PrinterTexto PosXX, PosLinea, SQLMsg2
      PosLinea = PosLinea + 0.4
   End If
   If SQLMsg3 <> "" Then
      Printer.FontSize = 9
      PrinterTexto InicX, PosLinea, SQLMsg3
      If SQLMsg4 <> "" Then
         PosXX = LimiteAncho - Printer.TextWidth(SQLMsg4) - 2
         PrinterTexto PosXX, PosLinea, SQLMsg4
      End If
      PosLinea = PosLinea + 0.4
   End If
   Printer.FontBold = True
   If Len(Titulo_Documento) > 1 Then
      Printer.FontSize = 10
      PrinterTexto CentrarTexto(UCaseStrg(Titulo_Documento)), PosLinea, UCaseStrg(Titulo_Documento)
      PosLinea = PosLinea + 0.5
   End If
   PosLinea = PosLinea + 0.05
   Imprimir_Linea_H PosLinea, InicX, AnchoDeLinea, Gris, True
   PosLinea = PosLinea + 0.1
   Printer.FontName = TipoArialNarrow
   Printer.FontBold = False
   Printer.FontSize = PorteLetra
   Printer.FontName = LetraAnterior
End Sub

Public Sub Encabezado(X0 As Single, x1 As Single)
Dim Y0 As Single
Dim y1 As Single
Dim Son_Iguales  As Boolean
   Son_Iguales = False
   PosLinea = 0.05
   PorteLetra = Printer.FontSize
   LetraAnterior = Printer.FontName
   Printer.FontName = TipoTimes
   Printer.FontSize = 8
   If X0 < 0.5 Then X0 = 0.5
   If Printer.Orientation = 1 Then Y0 = 0.01 Else Y0 = 0.5
   y1 = Y0 + 2.2
   'If X1 >= LimiteAncho Then X1 = LimiteAncho - 0.1
   PrinterPaint LogoTipo, X0, Y0 + 0.1, 3, 1.4
   RutaDestino = RutaSistema & "\LOGOS\DISKCOVS.GIF"
   PrinterPaint RutaDestino, x1 - 1.5, Y0 + 0.15, 1.5, 0.55
   'MsgBox RutaDestino
   Printer.FontBold = True: Printer.FontSize = 7
   Printer.CurrentX = x1 - 3.5: Printer.CurrentY = Y0 + 0.2
   Printer.Print "Hora:"
   Printer.CurrentX = x1 - 3.5: Printer.CurrentY = Y0 + 0.5
   Printer.Print "Pagina No."
   Printer.CurrentX = x1 - 3.5: Printer.CurrentY = Y0 + 0.8
   Printer.Print "Fecha:"
   Printer.CurrentX = x1 - 3.5: Printer.CurrentY = Y0 + 1.1
   Printer.Print "Usuario:"
   Printer.FontItalic = False
   Printer.FontBold = False: Printer.FontSize = 7
   Printer.CurrentX = x1 - 2.6: Printer.CurrentY = Y0 + 0.2
   Printer.Print Format$(Time, "hh:mm:ss")
   Printer.CurrentX = x1 - 2.2: Printer.CurrentY = Y0 + 0.5
   Printer.Print Format$(Pagina, "0000")
   Printer.CurrentX = x1 - 2.7: Printer.CurrentY = Y0 + 0.8
   Printer.Print FechaStrgDias(Date)
   Printer.CurrentX = x1 - 2.5: Printer.CurrentY = Y0 + 1.1
   Printer.Print ULCase(NombreUsuario)
   Printer.ForeColor = Negro
   Printer.FontItalic = True
   PosLinea = Y0
   If UCaseStrg(Empresa) = UCaseStrg(NombreComercial) Then Son_Iguales = True
   Printer.FontSize = 10
   Printer.FontBold = True
   Printer.CurrentX = CentrarTextoEncab(Empresa, X0, x1)
   Printer.CurrentY = PosLinea
   Printer.Print Empresa
   Printer.FontBold = False
   PosLinea = PosLinea + 0.5
   If (Son_Iguales = False) And NombreComercial <> Ninguno Then
      Printer.CurrentX = CentrarTextoEncab(NombreComercial, X0, x1)
      Printer.CurrentY = PosLinea
      Printer.Print NombreComercial
      PosLinea = PosLinea + 0.5
   End If '
   Printer.FontSize = 8
   Printer.FontBold = False
   Cadena = ULCase(NombreCiudad & ", " & Direccion & ". Teléfono: " & Telefono1 & ".")
   Printer.CurrentX = CentrarTextoEncab(Cadena, X0, x1)
   Printer.CurrentY = PosLinea
   Printer.Print Cadena
   PosLinea = PosLinea + 0.4
   If TrimStrg(MensajeEncabData) <> "" Then
      Printer.FontBold = True
      Printer.FontSize = 10
      PrinterTexto CentrarTextoEncab(MensajeEncabData, X0, x1), PosLinea, MensajeEncabData
      PosLinea = PosLinea + 0.5
      'MsgBox MensajeEncabData
   End If
   PosLinea = PosLinea + 0.1
   Printer.FontBold = False
   Pagina = Pagina + 1
   Printer.FontSize = PorteLetra
   Printer.FontName = LetraAnterior
   '  MsgBox PosLinea
   If PosLinea <= Y0 + 1.5 Then PosLinea = Y0 + 1.5
End Sub

Public Sub Encabezados()
Dim InicX As Single
Dim InicY As Single
Encabezado Ancho(0), AnchoPapel
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = True
If TrimStrg(SQLMsg1) <> "" Then
   Printer.FontSize = 13
   PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.5
End If
If TrimStrg(SQLMsg2) <> "" Then
   Printer.FontSize = 9
   PrinterTexto CentrarTexto(SQLMsg2), PosLinea, SQLMsg2
   PosLinea = PosLinea + 0.4
End If
If TrimStrg(SQLMsg3) <> "" Then
   Printer.FontSize = 8
   PrinterTexto Ancho(0), PosLinea, SQLMsg3
   PosLinea = PosLinea + 0.4
End If
' MsgBox "....................."
Printer.FontBold = False
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub EncabezadoSimple(X0 As Single, x1 As Single)
Dim Y0 As Single
Dim y1 As Single
   PosLinea = 0.1
   PorteLetra = Printer.FontSize
   LetraAnterior = Printer.FontName
   Printer.FontName = TipoTimes
   Printer.FontSize = 8
   If X0 < 0.5 Then X0 = 0.5
   Y0 = 1: y1 = Y0 + 1.5
   If x1 > LimiteAncho Then x1 = LimiteAncho - 0.1
   PrinterPaint LogoTipo, X0, Y0, 4, 2
   Printer.FontBold = True
   Printer.ForeColor = Gris
   Printer.FontItalic = True
   If UCaseStrg(Empresa) <> UCaseStrg(NombreComercial) Then
      Printer.FontSize = 16
      PrinterTexto CentrarTextoEncab(UCaseStrg(Empresa), X0, x1) + 0.025, Y0, UCaseStrg(Empresa)
      Y0 = Y0 + 0.025
      Printer.ForeColor = Negro
      PrinterTexto CentrarTextoEncab(UCaseStrg(Empresa), X0, x1), Y0, UCaseStrg(Empresa)
      Y0 = Y0 + 0.75
      Printer.FontSize = 14
      PrinterTexto CentrarTextoEncab(NombreComercial, X0, x1), Y0, NombreComercial
      Y0 = Y0 + 0.6
      Printer.FontSize = 8: Printer.FontItalic = False
      Cadena = Direccion & ". Teléfono: " & Telefono1 & "."
      PrinterTexto CentrarTextoEncab(Cadena, X0, x1), Y0, Cadena
   Else
      Printer.FontSize = 18
      PrinterTexto CentrarTextoEncab(UCaseStrg(Empresa), X0, x1) + 0.025, Y0, UCaseStrg(Empresa)
      Y0 = Y0 + 0.025
      Printer.ForeColor = Negro
      PrinterTexto CentrarTextoEncab(UCaseStrg(Empresa), X0, x1), Y0, UCaseStrg(Empresa)
      Y0 = Y0 + 0.75
   End If
   Printer.FontBold = True
   Printer.FontSize = 10
   Y0 = Y0 + 0.4
   Cadena = NombreCiudad & " - " & NombrePais
   PrinterTexto CentrarTextoEncab(Cadena, X0, x1), Y0, Cadena
   Y0 = Y0 + 0.5
   If MensajeEncabData <> "" Then
      Printer.FontSize = 16
      PrinterTexto CentrarTextoEncab(MensajeEncabData, X0, x1), Y0, MensajeEncabData
      Y0 = Y0 + 0.7
   End If
   If SQLMsg1 <> "" Then
      Printer.FontSize = 15
      PrinterTexto CentrarTexto(SQLMsg1), Y0, SQLMsg1
      Y0 = Y0 + 0.5
   End If
   If SQLMsg2 <> "" Then
      Printer.FontSize = 12
      PrinterTexto CentrarTexto(SQLMsg2), Y0, SQLMsg2
      Y0 = Y0 + 0.5
   End If
   Printer.FontBold = False
   If SQLMsg3 <> "" Then
      Printer.FontSize = 10
      PrinterTexto X0, Y0, SQLMsg3
      Y0 = Y0 + 0.5
   End If
   PosLinea = Y0
   Printer.FontSize = PorteLetra
   Printer.FontName = LetraAnterior
End Sub

''Private Function EstadoWinSock(TCPSocket As Winsock) As String
''    Select Case TCPSocket.State
''    Case 0: EstadoWinSock = "Cerrado"
''    Case 1: EstadoWinSock = "Abierto"
''    Case 2: EstadoWinSock = "Escuchando"
''    Case 3: EstadoWinSock = "Conexión pendiente..."
''    Case 4: EstadoWinSock = "Resolviendo host..."
''    Case 5: EstadoWinSock = "Host resuelto"
''    Case 6: EstadoWinSock = "Conectando..."
''    Case 7: EstadoWinSock = "Conectado"
''    Case 8: EstadoWinSock = "Cerrando..."
''    Case 9: EstadoWinSock = "ERROR !!"
''    Case Else
''            EstadoWinSock = "Desconocido (" & TCPSocket.State & ")"
''    End Select
''End Function

Public Function Digito_Verificador(NumeroRUC As String) As String
   RatonReloj

  'SP que determinar que tipo de contribuyente es y el codigo si es pasaporte
   Digito_Verificador_SP NumeroRUC
   If Tipo_RUC_CI.Tipo_Beneficiario <> "R" And Len(Tipo_RUC_CI.RUC_CI) = 13 Then
      If Ping_IP("srienlinea.sri.gob.ec") And UCase(GetUrlSource(urlEsUnRUC & Tipo_RUC_CI.RUC_CI)) = "TRUE" Then
         Tipo_RUC_CI.Tipo_Beneficiario = "R"
         Tipo_RUC_CI.Codigo_RUC_CI = MidStrg(Tipo_RUC_CI.RUC_CI, 1, 10)
         Tipo_RUC_CI.Digito_Verificador = MidStrg(Tipo_RUC_CI.RUC_CI, 10, 1)
      End If
   End If
   TipoBenef = Tipo_RUC_CI.Tipo_Beneficiario
   If Tipo_RUC_CI.Tipo_Beneficiario = "R" Then Tipo_Contribuyente_SP_MySQL Tipo_RUC_CI.RUC_CI, Tipo_RUC_CI.MicroEmpresa, Tipo_RUC_CI.AgenteRetencion
   RatonNormal
   Digito_Verificador = Tipo_RUC_CI.Digito_Verificador
End Function

'''Public Sub Buscar_Datos(DGBusq As DataGrid, _
'''                        AdoBusq As Adodc)
'''  RatonReloj
'''  TipoDatoBusq = 0
'''  SQLBusq = AdoBusq.RecordSource
'''  If DGBusq.Col > 0 Then
'''     CampoBusqueda = DGBusq.Columns(DGBusq.Col).Caption
'''     If AdoBusq.Recordset.RecordCount > 0 Then TipoDatoBusq = AdoBusq.Recordset.fields(CampoBusqueda).Type
'''  End If
'''  FBusqueda.Show 1
'''  With AdoBusq.Recordset
'''   If .RecordCount > 0 And TextoBusqueda <> "" Then
'''       RatonReloj
'''      .MoveFirst
'''      .Find (CampoBusqueda & TextoBusqueda)
'''       If .EOF Then
'''           MsgBox "No existe Datos que buscar"
'''          .MoveFirst
'''       End If
'''   Else
'''       MsgBox "No existe Datos que buscar"
'''   End If
'''  End With
'''  RatonNormal
'''  DGBusq.SetFocus
'''End Sub

Public Sub Imprimir_Documentos(NombreFile As String, _
                               SizeLetra As Integer, _
                               IniX As Single, _
                               IniY As Single, _
                               AnchoDeLinea As Single, _
                               TipoComprobante As String, _
                               Optional SQLx As String, _
                               Optional tipoDeLetra As String)
Dim NumFile As Integer
Dim NumPos As Long
Dim IE As Long
Dim KE As Long
Dim JE As Long
Dim No_Soc As Byte

Dim RutaGeneraFile As String
Dim Caracter As String
Dim Texto As String
Dim LineFormato As String
Dim LineTexto As String
Dim LineTextoTemp As String
Dim LineaDeMalla As String
Dim InsertarLinea As String
Dim CampoTexto As String
Dim ValorBool As String
Dim ElNoDias As String
Dim ElNoAnio As String
Dim ElTotal As String
Dim ElValor As String
Dim ElValorUnit As String
Dim DifBlancos As Single
Dim DimBlancos As Single
Dim XIniX As Currency
Dim XIniY As Currency
Dim CadFormato As String

Dim AdoDBx As ADODB.Recordset
On Error GoTo Errorhandler
ReDim Ancho(10) As Single
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = UCaseStrg(TipoComprobante)
Bandera = False
InsertarLinea = Producto
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj

InicioX = IniX: InicioY = IniY
Ancho(0) = IniX
LineaDeMalla = ""
If tipoDeLetra <> "" Then
   Printer.FontName = tipoDeLetra
Else
   Printer.FontName = TipoCourierNew
End If
Printer.FontSize = SizeLetra

If Len(SQLx) > 0 Then
 ' MsgBox SQLx
   Select_AdoDB AdoDBx, SQLx
   Ancho_Recordset InicioX, AdoDBx, SizeLetra, TipoCourierNew, 1, True
   With AdoDBx
    If .RecordCount > 0 Then
       .MoveFirst
        ReDim TipoC(.fields.Count - 1) As Campos_Tabla
        For I = 0 To .fields.Count - 1
            TipoC(I).Campo = .fields(I).Name
            TipoC(I).Ancho = AnchoTipoCampoTexto(.fields(I))
        Next I
        Do While Not .EOF
        For I = 0 To .fields.Count - 1
            Select Case .fields(I).Type
              Case TadByte, TadInteger, TadLong
                   LineaDeMalla = LineaDeMalla & SetearBlancos(.fields(I), TipoC(I).Ancho, 0, True, FAConLineas)
              Case TadSingle, TadDouble, TadCurrency
                   LineaDeMalla = LineaDeMalla & SetearBlancos(.fields(I), TipoC(I).Ancho, 0, True, FAConLineas, True)
              Case TadBoolean
                   ValorBool = "Si"
                   If .fields(I) = 0 Then ValorBool = "No"
                   LineaDeMalla = LineaDeMalla & " " & SetearBlancos(ValorBool, TipoC(I).Ancho, 0, True, FAConLineas)
              Case Else
                   LineaDeMalla = LineaDeMalla & " " & SetearBlancos(.fields(I), TipoC(I).Ancho, 0, False, FAConLineas)
            End Select
        Next I
        LineaDeMalla = LineaDeMalla & vbCrLf
       .MoveNext
        Loop
    End If
   End With
   'MsgBox LineaDeMalla
Else
   Escala_Centimetro 1, TipoCourierNew, SizeLetra, , True
End If
'Ejecutamos la consulta
ElTotal = Cambio_Letras(Total)
ElValor = Cambio_Letras(Valor)
ElValorUnit = Cambio_Letras(ValorUnit, True)
ElNoAnio = Cambio_Letras(NoAnio, True, True)
ElNoDias = Cambio_Letras(NoDias, True, True)
Pagina = 1
RutaGeneraFile = RutaSistema & "\DOCUMENT\" & NombreFile
NumFile = FreeFile
Texto = ""
PosLinea = IniY
'MsgBox RutaGeneraFile
Open RutaGeneraFile For Input As #NumFile ' Abre el archivo.
  Do While Not EOF(NumFile)
     Texto = Texto & Input(1, #NumFile)
  Loop
Close #NumFile

J = Len(Texto)
I = 1
LineTexto = ""
Do While I < J
   Evaluar = False
   Caracter = MidStrg(Texto, I, 1)
   CampoTexto = MidStrg(Texto, I, 3)
   Select Case CampoTexto
     Case "[A]": LineTexto = LineTexto & Carta_Porte: Evaluar = True
     Case "[a]": LineTexto = LineTexto & ElNoAnio: Evaluar = True
     Case "[B]": LineTexto = LineTexto & Beneficiario: Evaluar = True
     Case "[c]": LineTexto = LineTexto & NombreCiudad: Evaluar = True
     Case "[C]": LineTexto = LineTexto & NombreCliente: Evaluar = True
     Case "[d]": LineTexto = LineTexto & ElNoDias: Evaluar = True
     Case "[D]": LineTexto = LineTexto & DirCliente: Evaluar = True
     Case "[E]": LineTexto = LineTexto & Empresa: Evaluar = True
     Case "[e]": LineTexto = LineTexto & Institucion1 & " " & Institucion2: Evaluar = True
     Case "[f]": LineTexto = LineTexto & MesesLetras(NoMeses): Evaluar = True
     Case "[F]": LineTexto = LineTexto & FechaStrgCiudad(Mifecha): Evaluar = True
     Case "[I]": LineTexto = LineTexto & CICliente: Evaluar = True
     Case "[k]": LineTexto = LineTexto & Unidad: Evaluar = True
     Case "[K]": LineTexto = LineTexto & Cod_Bodega: Evaluar = True
     Case "[L]": LineTexto = LineTexto & LineasDeTexto: Evaluar = True
     Case "[m]": LineTexto = LineTexto & Producto: Evaluar = True
     Case "[M]": LineTexto = LineTexto & LineaDeMalla & vbCrLf: Evaluar = True
     Case "[n]": LineTexto = LineTexto & ElValorUnit: Evaluar = True
     Case "[N]": LineTexto = LineTexto & Format$(ValorUnit, "#,##0"): Evaluar = True
     Case "[p]": LineTexto = LineTexto & CStr(Interes * 100) & "%": Evaluar = True
     Case "[P]": LineTexto = LineTexto & CStr(NoDias): Evaluar = True
     Case "[S]": LineTexto = LineTexto & Format$(Factura_No, "000000"): Evaluar = True
     Case "[t]": LineTexto = LineTexto & ElTotal: Evaluar = True
     Case "[T]"
                 If (Total - Fix(Total)) > 0 Then
                    LineTexto = LineTexto & Format$(Total, "#,##0.00")
                 Else
                    LineTexto = LineTexto & Format$(Total, "#,##0")
                 End If
                 Evaluar = True
     Case "[v]": LineTexto = LineTexto & ElValor: Evaluar = True
     Case "[V]"
                 If (Valor - Fix(Valor)) > 0 Then
                    LineTexto = LineTexto & Format$(Valor, "#,##0.00")
                 Else
                    LineTexto = LineTexto & Format$(Valor, "#,##0")
                 End If
                 Evaluar = True
     Case "[y]": LineTexto = LineTexto & Codigo2: Evaluar = True
     Case Else:  LineTexto = LineTexto & Caracter
   End Select
   'MsgBox LineTexto
   If Evaluar Then I = I + 3 Else I = I + 1
Loop
' MsgBox LineTexto
' MsgBox TipoComprobante
Encabezado_Documento IniX, IniY, AnchoDeLinea, TipoComprobante
' Encabezados
' PrinterPaint LogoTipo, IniX, IniY, 2, 1.5
' PosLinea = IniY + 1.6
   If tipoDeLetra <> "" Then
      Printer.FontName = tipoDeLetra
   Else
      Printer.FontName = TipoCourierNew
   End If
Printer.FontSize = SizeLetra
Texto = LineTexto
J = Len(Texto)
I = 1
 'MsgBox Texto
LineTexto = ""
LineFormato = ""
Do While I < J
   Caracter = MidStrg(Texto, I, 1)
   CampoTexto = MidStrg(Texto, I, 3)
   LineTexto = LineTexto & Caracter
   
   LineTexto = Replace(LineTexto, vbCr, "")
   LineTexto = Replace(LineTexto, vbLf, "")
   'MsgBox "(" & LineTexto & ")"
   If Caracter = "[" Then LineFormato = LineFormato = "[XX"
   If Caracter = "]" Then LineFormato = LineFormato = "X]"
   If LineTexto = "" Then
      PosLinea = PosLinea + Printer.TextHeight("H")
      LineTexto = ""
      LineFormato = ""
   ElseIf LineTexto = vbCrLf Then
      PosLinea = PosLinea + Printer.TextHeight("H")
      LineTexto = ""
      LineFormato = ""
   ElseIf Printer.TextWidth(LineTexto) > AnchoDeLinea Or Asc(Caracter) = 13 Then
      If (Printer.TextWidth(LineTexto) - Len(LineFormato)) > AnchoDeLinea Then
         K = Len(LineTexto)
         If K > 0 Then
            Do
               K = K - 1
               I = I - 1
            Loop Until K < 2 Or MidStrg(LineTexto, K, 1) = " "
            LineTexto = MidStrg(LineTexto, 1, K)
         End If
         DimBlancos = Printer.TextWidth(" ")
         NumFacturas = (AnchoDeLinea - Printer.TextWidth(LineTexto)) / DimBlancos
         IE = 1
         LineTextoTemp = ""
         Do While NumFacturas > 1
            'Caracter = MidStrg(LineTexto, IE, 1)
            'MsgBox Caracter
            If MidStrg(LineTexto, IE, 1) = " " Then
               LineTextoTemp = LineTextoTemp & "  "
               NumFacturas = NumFacturas - 1
            Else
               LineTextoTemp = LineTextoTemp & MidStrg(LineTexto, IE, 1)
            End If
            'MsgBox NumFacturas & " (" & IE & ")" & vbCrLf & LineTextoTemp & vbCrLf & LineTexto
            IE = IE + 1
            If IE > Len(LineTexto) Then NumFacturas = 0
         Loop
         LineTextoTemp = LineTextoTemp & MidStrg(LineTexto, IE, Len(LineTexto))
         LineTexto = LineTextoTemp
      End If
     'Imprimir Linea con formato
      XIniX = IniX
      For KE = 1 To Len(LineTexto)
          CadFormato = ""
          If MidStrg(LineTexto, KE, 1) = "[" Then
             JE = KE
             Do While MidStrg(LineTexto, JE, 1) <> "]": JE = JE + 1: Loop
             CadFormato = MidStrg(LineTexto, KE, JE - KE + 1)
             KE = JE + 1
          End If
          'If CadFormato <> "" Then MsgBox LineTexto & vbCrLf & CadFormato & vbCrLf & KE
          Select Case CadFormato
            Case "[TN]": Printer.FontSize = SizeLetra
                         Printer.FontBold = False
                         Printer.FontItalic = False
                         Printer.FontUnderline = False
            Case "[F9]": Printer.FontSize = 9
            Case "[F10]": Printer.FontSize = 10
            Case "[F11]": Printer.FontSize = 11
            Case "[F12]": Printer.FontSize = 12
            Case "[F13]": Printer.FontSize = 13
            Case "[F14]": Printer.FontSize = 14
            Case "[F16]": Printer.FontSize = 16
            Case "[FN]": Printer.FontBold = True
            Case "[FC]": Printer.FontItalic = True
            Case "[FS]": Printer.FontUnderline = True
            Case "[NL]": Printer.NewPage
                         PosLinea = IniY + 2
            Case Else
                 Printer.CurrentX = XIniX
                 Printer.CurrentY = PosLinea
                 Printer.Print MidStrg(LineTexto, KE, 1)
                 XIniX = XIniX + Printer.TextWidth(MidStrg(LineTexto, KE, 1))
          End Select
      Next KE
      PosLinea = PosLinea + Printer.TextHeight("H")
      LineTexto = ""
      LineFormato = ""
'''      If Asc(Caracter) = 13 Then
'''         'MsgBox "13"
'''       '  I = I + 1
'''      End If
   End If
   If PosLinea >= LimiteAlto Then
      Printer.NewPage
      PosLinea = IniY + 2
      LineTexto = ""
   End If
   I = I + 1
Loop
Producto = InsertarLinea
Printer.EndDoc
RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub Imprimir_Liquidacion(SQLx As String, ReImp As Boolean)
Dim AdoDBx As ADODB.Recordset
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRIMIR LIQUIDACION DE CREDITO"
Bandera = False
SetPrinters.Show 1
Select_AdoDB AdoDBx, SQLx
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
Ancho_Recordset InicioX, AdoDBx, 9, TipoArial, 1, True
Pagina = 1
'MsgBox SQLx
With AdoDBx
 If .RecordCount > 0 Then
    .MoveFirst
     Total = 0                           'Total de Joyas (Capital)
     Mifecha = .fields("Fecha")
     NombreCliente = .fields("Cliente")
     DirCliente = .fields("Direccion")
     TotalCapital = .fields("Avaluo")
     Interes = .fields("Porc_Int") / 100
     NoDias = .fields("Plazo")
     Contador = 2
     Producto = .fields("Operacion") & " - " & .fields("Trans_No")
     Do While Not .EOF
        Total = Total + Fix(.fields("Valor"))
        Contador = Contador + 1
       .MoveNext
     Loop
    .MoveFirst
     TotalCredito = Total
     If 0 < TotalCapital And TotalCapital < Total Then Total = TotalCapital
     Total_Interes = Redondear(((Total * Interes) / 360) * NoDias, 2) ' Interes
     Saldo = Redondear(((Total * 0.005) / 360) * NoDias, 2)           ' Solca 0.5% y Tesoro Nacional 0.5%
     Total_Comision = Redondear((Total * 0.05), 2)                    ' Comision
     Total_IVA = Redondear(Total_Comision * Porc_IVA, 2)              ' I.V.A.
     Total_Desc = Total_Comision + Total_IVA                      ' Descuento
     Total_Factura = Total - Total_Comision - Total_IVA - Saldo - Saldo ' Total a Recibir
     Total_SubTotal = Total + Total_Interes                       ' Total a pagar al vencimiento
     Cadena2 = Cambio_Letras(Total_Factura) & " Dolares."
    .MoveFirst
 End If
End With
Si_No = True
Volver_Imp:
RatonReloj
'Iniciamos la impresion
Printer.FontBold = True
Printer.FontSize = 10
PosLinea = 0.1
PrinterTexto CentrarTexto(UCaseStrg(Empresa)), PosLinea, UCaseStrg(Empresa)
Printer.FontSize = 8
If ReImp Then PrinterTexto 16.5, PosLinea, "REIMPRESION"
PosLinea = PosLinea + 0.5
PrinterTexto CentrarTexto(RUC), PosLinea, RUC
PosLinea = PosLinea + 0.35
PrinterTexto CentrarTexto(Direccion), PosLinea, Direccion
PosLinea = 1.5
Printer.FontSize = 9
PrinterTexto CentrarTexto("LIQUIDACION DE OPERACION"), PosLinea, "LIQUIDACION DE OPERACION"
PosLinea = PosLinea + 1
PrinterTexto 1, PosLinea, "LUGAR"
PrinterTexto 12, PosLinea, "FECHA"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "CLIENTE"
PrinterTexto 12, PosLinea, "CEDULA/R.U.C."
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "PLAZO (días)"
PrinterTexto 5.5, PosLinea, "VENCIMIENTO"
PrinterTexto 12, PosLinea, "OPERACION"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "FORMA DE PAGO"
PrinterTexto 12, PosLinea, "PESO TOTAL (gr.)"
PosLinea = PosLinea + 0.6
Imprimir_Linea_H PosLinea, 1, 19
PosLinea = PosLinea + 0.05
PrinterTexto CentrarTexto("LIQUIDACION DE CREDITO"), PosLinea, "LIQUIDACION DE CREDITO"
PosLinea = PosLinea + 1
PrinterTexto 1, PosLinea, "VALOR DEL CREDITO"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "IMPUESTO SOLCA"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "TESORO NACIONAL"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "INTERES POR PAGAR"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "TOTAL A PAGAR AL VENCIMIENTO"
PosLinea = PosLinea + 0.6
Imprimir_Linea_H PosLinea, 1, 19
PosLinea = PosLinea + 0.05
PrinterTexto CentrarTexto("DETALLE DE LAS JOYAS ENTREGADAS EN CUSTODIA"), PosLinea, "DETALLE DE LAS JOYAS ENTREGADAS EN CUSTODIA"
PosLinea = PosLinea + 1
Printer.FontUnderline = True
PrinterTexto 1, PosLinea, "CANTIDAD"
PrinterTexto 3, PosLinea, "DESCRIPCION"
PrinterTexto 12, PosLinea, "PESO(gr)"
PrinterTexto 14.5, PosLinea, "QUILATAJE"
PrinterTexto 17, PosLinea, "VALOR"
Printer.FontUnderline = False
PosLinea = PosLinea + (Contador * 0.5)
Imprimir_Linea_H PosLinea, 1, 19
PosLinea = PosLinea + 0.05
If Si_No Then
PrinterTexto 1, PosLinea, "COMISION DE CUSTODIA"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "IMPUESTO AL VALOR AGREGADO"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "TOTAL A DESCONTAR"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "TOTAL A RECIBIR"
PosLinea = PosLinea + 0.5
Imprimir_Linea_H PosLinea, 1, 19
PosLinea = PosLinea + 0.05
TextoProc = "Declaro aceptar las condiciones pactadas en esta operación, sin tener la posibilidad de " _
          & "presentar reclamo alguno"
PrinterTexto 1, PosLinea, TextoProc
PosLinea = PosLinea + 0.4
PrinterTexto 1, PosLinea, "una vez firmada la presente liquidacion."
PosLinea = PosLinea + 0.5
Imprimir_Linea_H PosLinea, 1, 19

   PosLinea = PosLinea + 0.05
   PrinterTexto 1, PosLinea, "RECIBI CONFORME"
   PosLinea = PosLinea + 0.5
   PrinterTexto 1, PosLinea, "CUSTODIA EN"
   PosLinea = PosLinea + 1.5

End If
Printer.FontBold = False
PosLinea = 1.5
With AdoDBx
 If .RecordCount > 0 Then
    'IMPRESION DE LOS RESULTADOS
    .MoveFirst
     'MsgBox .RecordCount
     PosLinea = PosLinea + 1
     PrinterTexto 3.5, PosLinea, UCaseStrg(NombreCiudad)
     PrinterTexto 13.5, PosLinea, .fields("Fecha")
     PosLinea = PosLinea + 0.5
     PrinterTexto 3.5, PosLinea, .fields("Cliente")
     PrinterTexto 15, PosLinea, .fields("CI_RUC")
     PosLinea = PosLinea + 0.5
     PrinterFields 3.5, PosLinea, .fields("Plazo")
     PrinterFields 8.5, PosLinea, .fields("Fecha_V")
     PrinterTexto 14.5, PosLinea, .fields("Operacion") & "-" & CStr(.fields("Trans_No"))
     PosLinea = PosLinea + 0.5
     PrinterTexto 4, PosLinea, .fields("Forma_Pago")
     PrinterFields 15, PosLinea, .fields("Peso_Total")
     PosLinea = PosLinea + 1.65
     PrinterFields 14, PosLinea, .fields("Avaluo")
     PosLinea = PosLinea + 0.5
     PrinterVariables 14, PosLinea, Saldo
     PosLinea = PosLinea + 0.5
     PrinterVariables 14, PosLinea, Saldo
     PosLinea = PosLinea + 0.5
     PrinterVariables 14, PosLinea, Total_Interes
     PosLinea = PosLinea + 0.5
     PrinterVariables 14, PosLinea, Total_SubTotal
     PosLinea = PosLinea + 2.05
     Sumatoria = 0
     Do While Not .EOF
        PrinterFields 1, PosLinea, .fields("Cantidad")
        PrinterFields 3, PosLinea, .fields("Descripcion")
        PrinterFields 12.5, PosLinea, .fields("Peso")
        PrinterFields 14.5, PosLinea, .fields("Kilataje")
        PrinterFields 17, PosLinea, .fields("Valor")
        Sumatoria = Sumatoria + .fields("Valor")
        PosLinea = PosLinea + 0.55
       .MoveNext
     Loop
     Printer.Line (1, PosLinea)-(19, PosLinea), Negro
     PosLinea = PosLinea + 0.05
     PrinterTexto 14.5, PosLinea, "T O T A L"
     PrinterVariables 17, PosLinea, Sumatoria
     PosLinea = PosLinea + 0.55
    .MoveLast
     If Si_No Then
        PosLinea = PosLinea - 0.1
        PrinterVariables 14, PosLinea, Total_Comision
        PosLinea = PosLinea + 0.5
        PrinterVariables 14, PosLinea, Total_IVA
        PosLinea = PosLinea + 0.5
        PrinterVariables 14, PosLinea, Total_Desc
        PosLinea = PosLinea + 0.5
        PrinterVariables 14, PosLinea, Total_Factura
        PosLinea = PosLinea + 1.45
        
           PrinterTexto 4.5, PosLinea, Cadena2
           PosLinea = PosLinea + 0.5
           PrinterTexto 3.5, PosLinea, .fields("Operacion") & "-" & CStr(.fields("Trans_No"))
           PosLinea = PosLinea + 2
           PrinterTexto 1.5, PosLinea, .fields("Cliente")
           PrinterTexto 11, PosLinea, Empresa
           PosLinea = PosLinea + 0.45
           PrinterTexto 1.5, PosLinea, "C.I./R.U.C. " & .fields("CI_RUC")
        
     End If
 End If
End With
If Si_No Then
   Si_No = False
   Printer.NewPage
   GoTo Volver_Imp
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

Public Sub Imprimir_Cancelacion_Liquidacion(SQLx As String, ReImp As Boolean)
Dim AdoDBx As ADODB.Recordset
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRIMIR CANCELACION DE CREDITO"
Bandera = False
SetPrinters.Show 1
Select_AdoDB AdoDBx, SQLx
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
InicioX = 0.5: InicioY = 0
Ancho_Recordset InicioX, AdoDBx, 9, TipoArial, 1, True
Pagina = 1
'MsgBox SQLx
With AdoDBx
 If .RecordCount > 0 Then
    .MoveFirst
     Total = 0                           'Total de Joyas (Capital)
     Mifecha = .fields("Fecha")
     NombreCliente = .fields("Cliente")
     DirCliente = .fields("Direccion")
     TotalCapital = .fields("Avaluo")
     Interes = .fields("Porc_Int") / 100
     NoDias = .fields("Plazo")
     Contador = 2
     Do While Not .EOF
        Total = Total + Fix(.fields("Valor"))
        Contador = Contador + 1
       .MoveNext
     Loop
    .MoveFirst
     TotalCredito = Total
     If 0 < TotalCapital And TotalCapital < Total Then Total = TotalCapital
     Total_Interes = Redondear(((Total * Interes) / 360) * NoDias, 2) ' Interes
     Saldo = Redondear(((Total * 0.005) / 360) * NoDias, 2)           ' Solca 0.5% y Tesoro Nacional 0.5%
     Total_Comision = Redondear((Total * 0.05), 2)                    ' Comision
     Total_IVA = Redondear(Total_Comision * Porc_IVA, 2)              ' I.V.A.
     Total_Desc = Total_Comision + Total_IVA                      ' Descuento
     Total_Factura = Total - Total_Comision - Total_IVA - Saldo - Saldo ' Total a Recibir
     Total_SubTotal = Total + Total_Interes                       ' Total a pagar al vencimiento
     Cadena2 = Cambio_Letras(Total_Factura) & " Dolares."
    .MoveFirst
 End If
End With
Si_No = True
Volver_Imp:
RatonReloj
'Iniciamos la impresion
Printer.FontBold = True
Printer.FontSize = 10
PosLinea = 0.1
PrinterTexto CentrarTexto(UCaseStrg(Empresa)), PosLinea, UCaseStrg(Empresa)
Printer.FontSize = 8
If ReImp Then PrinterTexto 16.5, PosLinea, "REIMPRESION"
PosLinea = PosLinea + 0.5
PrinterTexto CentrarTexto(RUC), PosLinea, RUC
PosLinea = PosLinea + 0.35
PrinterTexto CentrarTexto(Direccion), PosLinea, Direccion
PosLinea = 1.5
Printer.FontSize = 9
PrinterTexto CentrarTexto("CANCELACION DE OPERACION"), PosLinea, "CANCELACION DE OPERACION"

PosLinea = PosLinea + 1
PrinterTexto 1, PosLinea, "LUGAR"
PrinterTexto 12, PosLinea, "FECHA CANCELACION"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "CLIENTE"
PrinterTexto 12, PosLinea, "CEDULA/R.U.C."
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "PLAZO (días)"
PrinterTexto 5.5, PosLinea, "VENCIMIENTO"
PrinterTexto 12, PosLinea, "OPERACION"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "FORMA DE PAGO"
PrinterTexto 12, PosLinea, "PESO TOTAL (gr.)"
PosLinea = PosLinea + 0.6
Printer.Line (1, PosLinea)-(19, PosLinea), Negro
PosLinea = PosLinea + 0.05
PrinterTexto CentrarTexto("CANCELACION DE CREDITO"), PosLinea, "CANCELACION DE CREDITO"
PosLinea = PosLinea + 1
If Si_No Then
PrinterTexto 1, PosLinea, "VALOR DEL CREDITO"
PosLinea = PosLinea + 0.5
PrinterTexto 1, PosLinea, "INTERES POR PAGAR"
PosLinea = PosLinea + 0.5
End If
PrinterTexto 1, PosLinea, "TOTAL A PAGAR"
PosLinea = PosLinea + 0.6
Imprimir_Linea_H PosLinea, 1, 19
PosLinea = PosLinea + 0.05
PrinterTexto CentrarTexto("DETALLE DE LAS JOYAS ENTREGADAS EN CUSTODIA"), PosLinea, "DETALLE DE LAS JOYAS ENTREGADAS EN CUSTODIA"
PosLinea = PosLinea + 1
Printer.FontUnderline = True
PrinterTexto 1, PosLinea, "CANTIDAD"
PrinterTexto 3, PosLinea, "DESCRIPCION"
PrinterTexto 12, PosLinea, "PESO(gr)"
PrinterTexto 14.5, PosLinea, "QUILATAJE"
PrinterTexto 17, PosLinea, "VALOR"
Printer.FontUnderline = False
PosLinea = PosLinea + (Contador * 0.5)
Printer.Line (1, PosLinea)-(19, PosLinea), Negro
PosLinea = PosLinea + 0.05
If Si_No Then
PosLinea = PosLinea + 0.05
TextoProc = "Declaro aceptar las condiciones de esta operación, sin tener la posibilidad de " _
          & "presentar reclamo alguno"
PrinterTexto 1, PosLinea, TextoProc
PosLinea = PosLinea + 0.4
PrinterTexto 1, PosLinea, "una vez firmada la presente cancelacion."
PosLinea = PosLinea + 0.5
Imprimir_Linea_H PosLinea, 1, 19
End If
Printer.FontBold = False
PosLinea = 1.5
With AdoDBx
 If .RecordCount > 0 Then
    'IMPRESION DE LOS RESULTADOS
    .MoveFirst
     'MsgBox .RecordCount
     PosLinea = PosLinea + 1
     PrinterTexto 3.5, PosLinea, UCaseStrg(NombreCiudad)
     PrinterTexto 16, PosLinea, FechaSistema
     'PrinterTexto 16, PosLinea, .Fields("Fecha_C")
     PosLinea = PosLinea + 0.5
     PrinterTexto 3.5, PosLinea, .fields("Cliente")
     PrinterTexto 15, PosLinea, .fields("CI_RUC")
     PosLinea = PosLinea + 0.5
     PrinterFields 3.5, PosLinea, .fields("Plazo")
     PrinterFields 8.5, PosLinea, .fields("Fecha_V")
     PrinterTexto 14.5, PosLinea, .fields("Operacion") & "-" & CStr(.fields("Trans_No"))
     PosLinea = PosLinea + 0.5
     PrinterTexto 4, PosLinea, .fields("Forma_Pago")
     PrinterFields 15, PosLinea, .fields("Peso_Total")
     PosLinea = PosLinea + 1.65
     If Si_No Then
     PrinterFields 14, PosLinea, .fields("Avaluo")
     PosLinea = PosLinea + 0.5
     PrinterVariables 14, PosLinea, Total_Interes
     PosLinea = PosLinea + 0.5
     End If
     PrinterVariables 14, PosLinea, Total_SubTotal
     PosLinea = PosLinea + 2.05
     Sumatoria = 0
     Do While Not .EOF
        PrinterFields 1, PosLinea, .fields("Cantidad")
        PrinterFields 3, PosLinea, .fields("Descripcion")
        PrinterFields 12.5, PosLinea, .fields("Peso")
        PrinterFields 14.5, PosLinea, .fields("Kilataje")
        PrinterFields 17, PosLinea, .fields("Valor")
        Sumatoria = Sumatoria + .fields("Valor")
        PosLinea = PosLinea + 0.55
       .MoveNext
     Loop
     Imprimir_Linea_H PosLinea, 1, 19
     PosLinea = PosLinea + 0.05
     PrinterTexto 14.5, PosLinea, "T O T A L"
     PrinterVariables 17, PosLinea, Sumatoria
     PosLinea = PosLinea + 0.55
    .MoveLast
     If Si_No Then
        PosLinea = PosLinea + 2
        PrinterTexto 1.5, PosLinea, .fields("Cliente")
        PrinterTexto 11, PosLinea, Empresa
        PosLinea = PosLinea + 0.45
        PrinterTexto 1.5, PosLinea, "C.I./R.U.C. " & .fields("CI_RUC")
        PosLinea = PosLinea + 1
        PrinterTexto 1.5, PosLinea, "Recibo conforme y entera satisfacion las Joyas aqui detalladas."
        PosLinea = PosLinea + 2
        PrinterTexto 1.5, PosLinea, .fields("Cliente")
        PosLinea = PosLinea + 0.45
        PrinterTexto 1.5, PosLinea, "C.I./R.U.C. " & .fields("CI_RUC")
     End If
 End If
End With
If Si_No Then
   Si_No = False
   Printer.NewPage
   GoTo Volver_Imp
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

Public Function PictCampo_Width(PictPrint As Object, _
                                AdoTipo As ADODB.Field, _
                                Optional ImpLineaCero As Boolean) As Single
Dim DistCampo As Single
Dim ch As String
DistCampo = 0
StrgAncho = AnchoTipoCampo(AdoTipo)
StrgFormatoCampo = FormatoTipoCampo(AdoTipo)
If StrgFormatoCampo = Ninguno Then
   StrgFormatoCampo = " "
ElseIf StrgFormatoCampo = "0" Or StrgFormatoCampo = "0.00" Then
   If ImpLineaCero Then StrgFormatoCampo = "--" Else StrgFormatoCampo = " "
End If
If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
    AltoLetra = PictPrint.PDFTextHeight
    LongNumero = PictPrint.PDFGetStringWidth(StrgAncho)
    LongCampo = PictPrint.PDFGetStringWidth(StrgFormatoCampo)
Else
    AltoLetra = PictPrint.TextHeight("H")
    LongNumero = PictPrint.TextWidth(StrgAncho)
    LongCampo = PictPrint.TextWidth(StrgFormatoCampo)
End If
Select Case AdoTipo.Type
  Case TadByte, TadCurrency, TadInteger, TadLong, TadSingle, TadDouble
       DistCampo = LongNumero - LongCampo
       If DistCampo <= 0 Then DistCampo = 0.05
End Select
If DistCampo <= 0 Then DistCampo = 0.05
PictCampo_Width = DistCampo
End Function

Public Function PictVariable_Width(PictPrint As Object, _
                                   Variable, _
                                   Optional CantDecimales As Byte) As Single
Dim DistVar As Single
Dim IdCero As Byte
Dim ParteDecimal As Integer
Dim TextoDecimal As String
DistVar = 0
StrgFormatoVariable = " "
Select Case VarType(Variable)
  Case vbBoolean: If Variable Then StrgFormatoVariable = "X"
  Case vbByte, vbInteger, vbLong
       StrgFormatoVariable = Str(Variable)
  Case vbSingle
       StrgFormatoVariable = Format$(Variable, "##0.00%")
  Case vbDouble, vbCurrency
       If CantDecimales > 2 Then
          TextoDecimal = CStr(Variable - Int(Variable))
          TextoDecimal = MidStrg(TextoDecimal, 3, Len(TextoDecimal))
          If Len(TextoDecimal) > 4 Then TextoDecimal = MidStrg(TextoDecimal, 1, 4)
          If Len(TextoDecimal) < 2 Then TextoDecimal = "00"
          StrgFormatoVariable = Format$(Variable, "#,##0." & String(Len(TextoDecimal), "0"))
       Else
          StrgFormatoVariable = Format$(Variable, "#,##0.00")
       End If
  Case vbString:  StrgFormatoVariable = Variable
  Case Else:      StrgFormatoVariable = Variable
End Select
StrgAncho = Ancho_Tipo_Variable(Variable)
If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
   AltoLetra = PictPrint.PDFTextHeight
   LongNumero = PictPrint.PDFGetTextWidth(StrgAncho)
   LongVariable = PictPrint.PDFGetTextWidth(StrgFormatoVariable)
Else
   AltoLetra = PictPrint.TextHeight(StrgAncho)
   LongNumero = PictPrint.TextWidth(StrgAncho)
   LongVariable = PictPrint.TextWidth(StrgFormatoVariable)
End If
Select Case VarType(Variable)
  Case vbByte, vbDouble, vbInteger, vbLong, vbSingle, vbCurrency
       DistVar = LongNumero + 0.05 - LongVariable
       If DistVar <= 0 Then DistVar = 0.05
       If Not TypeOf PictPrint Is mjwPDF Then
          If PictPrint.FontBold Then DistVar = DistVar - 0.04
       End If
       'If Val(StrgFormatoVariable) = 0 Then StrgFormatoVariable = "0"
  Case Else
       DistVar = 0
End Select
If DistVar <= 0 Then DistVar = 0.05
PictVariable_Width = DistVar
End Function

Public Sub PictPrint_Fields(PictPrint As Object, _
                            Xo As Single, _
                            Yo As Single, _
                            AdoTipo As ADODB.Field, _
                            Optional JustDer As Boolean, _
                            Optional anchoTexto As Single, _
                            Optional ImpLineaCero As Boolean)
Dim Xo1 As Single
    If (Xo > 0) And (Yo > 0) Then
       Xo1 = Xo
       Distancia = PictCampo_Width(PictPrint, AdoTipo, ImpLineaCero)
       If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
          If JustDer Then Xo1 = Xo + anchoTexto - PictPrint.PDFGetStringWidth(StrgFormatoCampo) + 0.05
          PictPrint.PDFTextOut StrgFormatoCampo, CDbl(Xo1), CDbl(Yo)
       Else
          If JustDer Then Xo1 = Xo + anchoTexto - PictPrint.TextWidth(StrgFormatoCampo) + 0.05
          PictPrint.CurrentY = Yo
          PictPrint.CurrentX = Xo1
          PictPrint.Print StrgFormatoCampo
       End If
      'MsgBox "Variable " & StrgFormatoCampo
    End If
End Sub

Public Sub PictPrint_Cuadro_Linea(PictPrint As Object, _
                                  Xo As Single, _
                                  Yo As Single, _
                                  Xf As Single, _
                                  Yf As Single, _
                                  Optional color As Long, _
                                  Optional BF As String)
Dim PXo As Double
Dim PYo As Double
Dim PXf As Double
Dim PYf As Double

    PXo = Xo: PYo = Yo: PXf = Xf: PYf = Yf
    If (PXo > 0) And (PYo > 0) And (PXf > 0) And (PYf > 0) Then
       If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
          PictPrint.PDFSetLineWidth = 0.01
          PictPrint.PDFSetBorder = BORDER_ALL
          PictPrint.PDFSetFill = True
          If BF = "BF" Then
            'PXo = PXo - 0.2
             PYo = PYo - 0.9
             PYf = PYf - 0.9
             PictPrint.PDFSetLineColor = Negro
             PictPrint.PDFSetDrawColor = color
             PictPrint.PDFSetDrawMode = DRAW_DRAW
             PictPrint.PDFDrawRectangle PXo, PYo, PXf - PXo, PYf - PYo
          ElseIf BF = "B" Then
            'PXo = PXo - 0.2
             PYo = PYo - 0.9
             PYf = PYf - 0.9
             PictPrint.PDFSetLineColor = color
             PictPrint.PDFDrawRectangle PXo, PYo, PXf - PXo, PYf - PYo
          Else
            PictPrint.PDFSetLineColor = color
            If PXo = PXf Then
               PYo = PYo - 0.5
               PYf = PYf - 0.5
            End If

            If PYo = PYf Then
               PYo = PYo - 0.5
               PYf = PYf - 0.5
              'PXo = PXo - 0.2
'            Else
'               PYf = PYf - 0.4
            End If
             PictPrint.PDFDrawLine PXo, PYo, PXf, PYf
          End If
          PictPrint.PDFSetDrawMode = DRAW_NORMAL
          PictPrint.PDFSetTextColor = Negro
       Else
          If color = 0 Then color = Negro
          If BF = "BF" Then
             PictPrint.Line (PXo, PYo)-(PXf, PYf), color, BF
          ElseIf BF = "B" Then
             PictPrint.Line (PXo, PYo)-(PXf, PYf), color, B
          Else
             PictPrint.Line (PXo, PYo)-(PXf, PYf), color
          End If
       End If
    End If
End Sub

Public Sub PictPrint_Estilo_Letra(PictPrint As Object, estiloLetra As PDFFontStl, Estado As Boolean)
    If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
       PictPrint.PDFSetFontStyle estiloLetra, Estado
    Else
       Select Case estiloLetra
        Case FONT_NORMAL
             PictPrint.FontBold = False
             PictPrint.FontUnderline = False
             PictPrint.FontItalic = False
        Case FONT_ITALIC: PictPrint.FontItalic = Estado
        Case FONT_BOLD: PictPrint.FontBold = Estado
        Case FONT_UNDERLINE: PictPrint.FontUnderline = Estado
       End Select
    End If
End Sub

Public Sub PictPrint_Tipo_Letra(PictPrint As Object, TipoLetra As String, Optional PorteLetra As Integer)
Dim PDFPorteLetra As Integer
    PDFPorteLetra = PorteLetra
    If PDFPorteLetra <= 0 Then PDFPorteLetra = 9
    If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
       Select Case TipoLetra
         Case TipoArial: PictPrint.PDFSetFont FONT_ARIAL, PDFPorteLetra, FONT_NORMAL
         Case TipoCourier: PictPrint.PDFSetFont FONT_COURIER, PDFPorteLetra, FONT_NORMAL
         Case TipoTimes: PictPrint.PDFSetFont FONT_TIMES, PDFPorteLetra, FONT_NORMAL
         Case TipoSymbol: PictPrint.PDFSetFont FONT_SYMBOL, PDFPorteLetra, FONT_NORMAL
         Case TipoHelvetica: PictPrint.PDFSetFont FONT_HELVETICA, PDFPorteLetra, FONT_NORMAL
         Case TipoTerminal: PictPrint.PDFSetFont FONT_ZAPFDINGBATS, PDFPorteLetra, FONT_NORMAL
         Case Else: PictPrint.PDFSetFont FONT_ARIAL, PDFPorteLetra, FONT_NORMAL
       End Select
    Else
       PictPrint.FontName = TipoLetra
    End If
End Sub

Public Sub PictPrint_Porte_Letra(PictPrint As Object, PorteLetra As Integer)
Dim PDFPorteLetra As Integer
    PDFPorteLetra = PorteLetra
    If PDFPorteLetra <= 0 Then PDFPorteLetra = 9
    If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
''       Select Case TipoLetra
''         Case TipoArial: PictPrint.PDFSetFont FONT_ARIAL, PDFPorteLetra, FONT_NORMAL
''         Case TipoCourier: PictPrint.PDFSetFont FONT_COURIER, PDFPorteLetra, FONT_NORMAL
''         Case TipoTimes: PictPrint.PDFSetFont FONT_TIMES, PDFPorteLetra, FONT_NORMAL
''         Case TipoSymbol: PictPrint.PDFSetFont FONT_SYMBOL, PDFPorteLetra, FONT_NORMAL
''         Case TipoHelvetica: PictPrint.PDFSetFont FONT_HELVETICA, PDFPorteLetra, FONT_NORMAL
''         Case TipoTerminal: PictPrint.PDFSetFont FONT_ZAPFDINGBATS, PDFPorteLetra, FONT_NORMAL
''         Case Else: PictPrint.PDFSetFont FONT_ARIAL, PDFPorteLetra, FONT_NORMAL
''       End Select
    Else
       PictPrint.FontSize = PDFPorteLetra
    End If
End Sub

Public Sub PictPrint_Color_Letra(PictPrint As Object, color As Long)
If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
   PictPrint.PDFSetTextColor = color
Else
   PictPrint.ForeColor = color
End If
End Sub

Public Sub PictPrint_Color_Fondo(PictPrint As Object, color As Long)
If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
   PictPrint.PDFSetDrawColor = color
Else
   PictPrint.FillColor = color
End If
End Sub

Public Sub PictPrint_Variables(PictPrint As Object, _
                               Xo As Single, _
                               Yo As Single, _
                               Variable, _
                               Optional JustDer As Boolean, _
                               Optional anchoTexto As Single, _
                               Optional ImpLineaCero As Boolean, _
                               Optional CantDecimales As Byte)
Dim NewFont As Long
Dim OldFont As Long
Dim IText As Integer
Dim SizeLetra As Integer
Dim PorteDeLaLetra As Single
Dim AnchoDelTexto As Single

If (Xo > 0) And (Yo > 0) Then
   Distancia = PictVariable_Width(PictPrint, Variable, CantDecimales)
   If StrgFormatoVariable = Ninguno Then
      StrgFormatoVariable = " "
   ElseIf StrgFormatoVariable = "0" Or StrgFormatoVariable = "0.00" Then
      If ImpLineaCero Then StrgFormatoVariable = "--" Else StrgFormatoVariable = " "
   End If
   If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
      If JustDer Then
         AnchoDelTexto = PictPrint.PDFGetTextWidth(StrgFormatoVariable)
         PictPrint.PDFTextOut StrgFormatoVariable, CDbl(Xo + anchoTexto - AnchoDelTexto + 0.05) - 0.99, CDbl(Yo) - 0.5
      Else
         PictPrint.PDFTextOut StrgFormatoVariable, CDbl(Xo) - 0.99, CDbl(Yo) - 0.5
      End If
   Else
      PictPrint.CurrentY = Yo
      If JustDer Then
         AnchoDelTexto = PictPrint.TextWidth(StrgFormatoVariable)
         PictPrint.CurrentX = Xo + anchoTexto - AnchoDelTexto + 0.05
      Else
         PictPrint.CurrentX = Xo
      End If
      PictPrint.Print StrgFormatoVariable
   End If
  'MsgBox "Variable " & StrgFormatoVariable
End If
End Sub

Public Sub PictPrint_Texto(PictPrint As Object, _
                           Xo As Single, _
                           Yo As Single, _
                           Texto As String, _
                           Optional JustDer As Boolean, _
                           Optional anchoTexto As Single, _
                           Optional CentrarElTexto As Boolean)
Dim Xo1 As Double
Dim Yo1 As Double
Dim AnchoText As Single
Dim Texto1 As String
Dim CantTexto As String
Dim Caracter As String
Dim PosTexto As Long
If ((Xo > 0) And (Yo > 0) And (Texto <> "")) Then
   Xo1 = CDbl(Xo)
   Yo1 = CDbl(Yo)
   Lineas_Impresas = 0
   If Texto = Ninguno Then Texto = " "
   If LimiteAncho <= 0 Then LimiteAncho = 20

   If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
      Xo1 = Xo1 - 0.99
      Yo1 = Yo1 - 0.55
      If Xo1 < 0 Then Xo1 = 0.01
      If Yo1 < 0 Then Yo1 = 0.01

      AnchoText = 0
      If JustDer Then
         AnchoText = PictPrint.PDFGetTextWidth(Texto)
         Xo1 = Xo1 + anchoTexto - AnchoText + 0.05
      End If
      If CentrarElTexto Then
         AnchoText = PictPrint.PDFGetTextWidth(Texto)
         If anchoTexto > 0 Then
            Xo1 = Xo1 + (anchoTexto / 2) - (AnchoText / 2)
         Else
            Xo1 = (PictPrint.PDFPageWidth / 2) - (AnchoText / 2)
         End If
      End If
      If anchoTexto > 0 Then
         Texto1 = Texto
         Do While Len(Texto1) > 0
            AnchoText = PictPrint.PDFGetTextWidth(Texto1)
            If AnchoText > anchoTexto Then
               CantTexto = ""
               PosTexto = 1
               AnchoText = 0
               Do While AnchoText <= anchoTexto
                  CantTexto = CantTexto & MidStrg(Texto1, PosTexto, 1)
                  AnchoText = PictPrint.PDFGetTextWidth(CantTexto)
                  PosTexto = PosTexto + 1
                  If PosTexto > Len(Texto1) Then AnchoText = anchoTexto + 1
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
               CantTexto = Texto1
            End If
            PictPrint.PDFTextOut CantTexto, Xo1, Yo1
            Lineas_Impresas = Lineas_Impresas + 1
            Yo1 = Yo1 + 0.3
            Yo = Yo + 0.3
            If Len(Texto1) >= Len(CantTexto) Then
               Texto1 = TrimStrg(MidStrg(Texto1, Len(CantTexto) + 1, Len(Texto1)))
            Else
               Texto1 = ""
            End If
        Loop
        'MsgBox Lineas_Impresas
        If Lineas_Impresas >= 1 Then Yo = Yo - 0.35
      Else
         'MsgBox Texto
         PictPrint.PDFTextOut Texto, Xo1, Yo1
         Lineas_Impresas = Lineas_Impresas + 1
      End If
    Else                                 ' Si la impresion en es Printer o Picture
      If JustDer Then Xo1 = Xo1 + anchoTexto - PictPrint.TextWidth(Texto) + 0.05
      If CentrarElTexto Then
         If anchoTexto > 0 Then
            Xo1 = Xo1 + (anchoTexto / 2) - (PictPrint.TextWidth(Texto) / 2)
         Else
            Xo1 = (LimiteAncho / 2) - (PictPrint.TextWidth(Texto) / 2)
         End If
      End If
      PictPrint.CurrentX = Xo1
      PictPrint.CurrentY = Yo1
      PictPrint.Print Texto
      Lineas_Impresas = Lineas_Impresas + 1
   End If
End If
End Sub

Public Function PictPrint_Texto_Multiple(PictPrint As Object, _
                                         Xo As Single, _
                                         Yo As Single, _
                                         Texto As String, _
                                         Optional anchoTexto As Single) As Single
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
Yo1 = Yo
If ((Xo > 0) And (Yo > 0) And (Texto <> "")) Then
   If LimiteAncho <= 0 Then LimiteAncho = 20
   If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
      AltoDeLetras = PictPrint.PDFTextHeight
      AnchoDeCar = PictPrint.PDFGetTextWidth("H")
   Else
      AltoDeLetras = PictPrint.TextHeight("H")
      AnchoDeCar = PictPrint.TextWidth("H")
   End If
   TextTemp = Texto
   TextAux = ""
   PosCar = 0
   Do While Len(TextTemp) > 1 And PosCar <= Len(TextTemp)
      PosCar = PosCar + 1
      Caracter = MidStrg(TextTemp, PosCar, 1)
      CaracterSig = MidStrg(TextTemp, PosCar + 1, 1)
      TextAux = TextAux & Caracter
      If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
         AnchoDeLetras = PictPrint.PDFGetTextWidth(TextAux)
      Else
         AnchoDeLetras = PictPrint.TextWidth(TextAux)
      End If
      If Caracter = vbCr Then
         cPrint.printTexto Xo, Yo1, TextAux
         Yo1 = Yo1 + AltoDeLetras
         TextTemp = TrimStrg(MidStrg(TextTemp, PosCar + 2, Len(TextTemp)))
         TextAux = ""
         PosCar = 0
      ElseIf AnchoDeLetras > anchoTexto Then
         If CaracterSig <> " " Then
            Do While MidStrg(TextAux, PosCar, 1) <> " " And PosCar >= 1
               PosCar = PosCar - 1
            Loop
         End If
         TextAux = TrimStrg(MidStrg(TextAux, 1, PosCar))
         TextAux1 = TextAux
         If anchoTexto >= Redondear(AnchoDeLetras, 2) Then
            CantBlanco = Redondear((anchoTexto - Redondear(AnchoDeLetras, 2)) / Redondear(AnchoDeCar, 2))
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
'''               & AnchoTexto & vbCrLf & Redondear(PictPrint.TextWidth(TextAux), 2) & vbCrLf _
'''               & TextAux & vbCrLf _
'''               & TextAux1 & vbCrLf
         cPrint.printTexto Xo, Yo1, TextAux
         Yo1 = Yo1 + AltoDeLetras
         TextTemp = TrimStrg(MidStrg(TextTemp, PosCar + 1, Len(TextTemp)))
         TextAux = ""
         PosCar = 0
      End If
   Loop
   If Len(TextAux) > 0 Then cPrint.printTexto Xo, Yo1, TextAux
   'MsgBox "Residuo: " & TextTemp & vbCrLf & TextAux
End If
PictPrint_Texto_Multiple = Yo1
End Function


Public Sub PictPrint_Nota_Materia(PictPrint As Object, _
                                  Xo As Single, _
                                  Yo As Single, _
                                  Nota_Materia As Variant, _
                                  Optional Cualitativa As Boolean, _
                                  Optional CDecimal As Byte, _
                                  Optional EsPreBasica As Boolean, _
                                  Optional Cualitativa2 As Boolean)
Dim Equivalente As String
Dim MiNota As Currency
Dim E As Byte
    'MsgBox Nota_Materia
    Equivalente = ""
    If Nota_Materia >= 0 Then
      'MiNota = Redondear(Nota_Materia, 2)
       
       MiNota = Redondear_2Dec(Nota_Materia)
       If MidStrg(FormatoLibreta, 1, 9) = "QUIMESTRE" Then
          If EsPreBasica Then
             For E = 0 To UBound(Equivalencias) - 1
                 If (Equivalencias(E).Desde <= MiNota) And (MiNota <= Equivalencias(E).Hasta) Then
                    'MsgBox MiNota
                    If Cualitativa Then
                       Equivalente = Equivalencias(E).Cualitativa
                    ElseIf Cualitativa2 Then
                       Equivalente = Equivalencias(E).Cualitativa2
                    Else
                       Equivalente = Equivalencias(E).Letras
                    End If
                    If Cualitativa And Cualitativa2 Then Equivalente = Equivalencias(E).Letras
                 End If
             Next E
          Else
             For E = 0 To UBound(Equivalencias) - 1
                 If (Equivalencias(E).Desde <= MiNota) And (MiNota <= Equivalencias(E).Hasta) Then
                    If Cualitativa Then
                       Equivalente = Equivalencias(E).Cualitativa
                    ElseIf Cualitativa2 Then
                       Equivalente = Equivalencias(E).Cualitativa2
                    Else
                       If CDecimal > 0 Then
                          Equivalente = Format$(Nota_Materia, "00." & String$(CDecimal, "0"))
                       Else
                          Equivalente = Format$(Nota_Materia, "00")
                       End If
                    End If
                    If Cualitativa And Cualitativa2 Then Equivalente = Equivalencias(E).Letras
                 End If
             Next E
          End If
          'MsgBox Equivalente & vbCrLf & Cualitativa & vbCrLf & Cualitativa2
        Else
          If Cualitativa Then
''             Select Case Nota_Materia
''               Case 0 To 11.49: Equivalente = "I"
''               Case 11.5 To 13.49: Equivalente = "R"
''               Case 13.5 To 15.49: Equivalente = "B"
''               Case 15.5 To 18.49: Equivalente = "MB"
''               Case 18.5 To 20: Equivalente = "S"
''             End Select
          Else
             If CDecimal > 0 Then
                Equivalente = Format$(Nota_Materia, "00." & String$(CDecimal, "0"))
             Else
                Equivalente = Format$(Nota_Materia, "00")
             End If
          End If
       End If
       If Nota_Materia = 0 Then
          If SinImprimir Then Equivalente = ""
          If Not ImpCeros Then Equivalente = ""
       End If
       cPrint.letraEstilo FONT_NORMAL, True
       If Nota_Materia < Nota_Rojo Then
          cPrint.letraEstilo FONT_UNDERLINE, True
          cPrint.letraEstilo FONT_BOLD, True
          If Len(Equivalente) = 1 Then
             cPrint.printTexto Xo + 0.09, Yo, Equivalente
          Else
             cPrint.printTexto Xo - 0.01, Yo, Equivalente
          End If
       End If
       If Len(Equivalente) = 1 Then
          cPrint.printTexto Xo + 0.1, Yo, Equivalente
       Else
          cPrint.printTexto Xo, Yo, Equivalente
       End If
       cPrint.letraEstilo FONT_UNDERLINE, False
       cPrint.letraEstilo FONT_NORMAL, True
       cPrint.colorDeLetra = vbBlack
    End If
End Sub

Sub Recordar_Tarea_Hora()
    Minutos = Minute(Time - TiempoTarea)
    Segundos = Second(Time - TiempoTarea)
    MiTiempo = CSng(Minutos & "." & Segundos)
    If MiTiempo >= 0.15 Then
       'MsgBox MiTiempo
       TiempoTarea = Time
       If CodigoUsuario <> "" And NumEmpresa <> "" Then
          'MsgBox UCaseStrg(NombreUsuario) & vbCrLf & vbCrLf & "INGRESE TAREA ACTUAL"
       End If
    End If
End Sub

Public Sub Imprimir_Clientes(DataCli As Adodc)
On Error GoTo Errorhandler
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION DE CLIENTES"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
RatonReloj
Pagina = 1
InicioX = 1: InicioY = 0
DataAnchoCampos 1, DataCli, 9, TipoCourier, 1, False
'Recolectamos los item de la factura a buscar
'MargIzq = 1: MargSup = 1:
PFil = 0
LimiteAlto = LimiteAlto - 1.5
With DataCli.Recordset
 If .RecordCount > 0 Then
     Encabezado 1, 19
     PosLinea = PosLinea + 0.1
     Imprimir_Linea_H PosLinea, 1, 19, Negro
    .MoveFirst
     Printer.FontName = TipoSansSerif
     Printer.FontSize = 8
     Do While Not .EOF
        PosLinea = PosLinea + 0.05
        Printer.FontBold = True
        PrinterTexto 1, PosLinea, "Cliente:"
        PrinterTexto 11, PosLinea, "CI_RUC:"
        PrinterTexto 15.5, PosLinea, "Codigo:"
        PrinterTexto 1, PosLinea + 0.36, "Grupo:"
        PrinterTexto 3.6, PosLinea + 0.36, "FAX:"
        PrinterTexto 6, PosLinea + 0.36, "Telefono(s):"
        PrinterTexto 12.2, PosLinea + 0.36, "Email:"
        PrinterTexto 1, PosLinea + 0.72, "Direccion:"
        PrinterTexto 12.2, PosLinea + 0.72, "Comercial:"
        Printer.FontBold = False
        PrinterFields 2.2, PosLinea, .fields("Cliente")
        PrinterFields 12.5, PosLinea, .fields("CI_RUC")
        PrinterFields 16.7, PosLinea, .fields("Codigo")
        PrinterFields 2.1, PosLinea + 0.36, .fields("Grupo")
        If Val(.fields("FAX")) <> 0 Then PrinterFields 4.4, PosLinea + 0.36, .fields("FAX")
        Codigo = ""
        If Len(.fields("Telefono")) > 1 And Val(.fields("Telefono")) <> 0 Then Codigo = Codigo & "/" & .fields("Telefono")
        If Len(.fields("TelefonoT")) > 1 And Val(.fields("TelefonoT")) <> 0 Then Codigo = Codigo & "/" & .fields("TelefonoT")
        If Len(.fields("Celular")) > 1 And Val(.fields("Celular")) <> 0 Then Codigo = Codigo & "/" & .fields("Celular")
        PrinterTexto 7.7, PosLinea + 0.36, Codigo
        PrinterFields 13.3, PosLinea + 0.36, .fields("Email")
        PrinterTexto 2.5, PosLinea + 0.72, .fields("Direccion") & " (" & .fields("DirNumero") & ")"
        PrinterTexto 14, PosLinea + 0.72, .fields("Representante")
        PosLinea = PosLinea + 1.09
        Imprimir_Linea_H PosLinea, 1, 19, Negro
        PosLinea = PosLinea + 0.05
        If PosLinea >= LimiteAlto Then
           Printer.NewPage
           Encabezado 1, 19
           PosLinea = PosLinea + 0.1
           Imprimir_Linea_H PosLinea, 1, 19, Negro
           Printer.FontName = TipoSansSerif
           Printer.FontSize = 8
        End If
       .MoveNext
     Loop
    .MoveFirst
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

Public Function VerdadFalso(Valor) As Boolean
 If Valor = 0 Then VerdadFalso = False Else VerdadFalso = True
End Function

Public Function ValorVerdadFalso(Valor As Boolean) As Integer
 If Valor = False Then ValorVerdadFalso = 0 Else ValorVerdadFalso = 1
End Function

Public Function CTNumero(ValorStr As String, Optional Decimales As Byte) As Currency
Dim Resul As Currency
Dim NumeroStr As String
    Resul = 0
    NumeroStr = TrimStrg(Replace(ValorStr, ",", ""))
    If IsNumeric(NumeroStr) Then
       If Decimales > 0 Then
          Resul = CCur(Format$(Val(NumeroStr), "##0." & String$(Decimales, "0")))
       Else
          Resul = CCur(Format$(Val(NumeroStr), "##0"))
       End If
    End If
    CTNumero = Resul
End Function

Public Function Cambio_CI_Caracter(CI As String) As String
Dim Cod_CI As String
Dim L1, L2, L3, L4, L5 As Byte
   If CI <> "" Then
      CI = Format$(Val(CI), String$(10, "0"))
      L1 = 65 + Val(MidStrg(CI, 1, 2)): If L1 >= 90 Then L1 = 65 + (L1 - 90)
      L2 = 65 + Val(MidStrg(CI, 3, 2)): If L2 >= 90 Then L2 = 65 + (L2 - 90)
      L3 = 65 + Val(MidStrg(CI, 5, 2)): If L3 >= 90 Then L3 = 65 + (L3 - 90)
      L4 = 65 + Val(MidStrg(CI, 7, 2)): If L4 >= 90 Then L4 = 65 + (L4 - 90)
      L5 = 65 + Val(MidStrg(CI, 9, 2)): If L5 >= 90 Then L5 = 65 + (L5 - 90)
      Cod_CI = Chr(L1) & Chr(L2) & Chr(L3) & Chr(L4) & Chr(L5)
   Else
      Cod_CI = "ABCDE"
   End If
   Cambio_CI_Caracter = UCaseStrg(Cod_CI)
End Function

Public Function Cambio_Usuario_Inicial(Usuario As String) As String
Dim Cod_CI As String
Dim Usuario1 As String
Dim L1 As Long
Dim L2 As Long
Dim L3 As Long

   Cod_CI = MidStrg(Usuario, 1, 1) & "."
   For L1 = 1 To Len(Usuario)
    If MidStrg(Usuario, L1, 1) = " " Then Cod_CI = Cod_CI & MidStrg(Usuario, L1 + 1, 1) & "."
   Next L1
   Cambio_Usuario_Inicial = UCaseStrg(TrimStrg(Cod_CI))
End Function

Public Function FechaDeZip(Fechas As String) As String
Dim FechaC As String
Dim FechaCar As String
  FechaC = FechaStrgCorta(Fechas)
  FechaCar = ""
  For I = 1 To Len(FechaC)
      If MidStrg(FechaC, I, 1) = "/" Then
         FechaCar = FechaCar & "_"
      Else
         FechaCar = FechaCar & MidStrg(FechaC, I, 1)
      End If
  Next I
  FechaDeZip = FechaCar
End Function

Public Sub Fecha_Del_AT(ATMes As String, ATAno As String)
  'Mes de creacion del XML
  'MsgBox ATMes
   Select Case ATMes
     Case "1er_Semestre"
          NumMeses = 6
          FechaInicial = "01/01/" & Format$(Val(ATAno), "0000")
          FechaMitad = "15/01/" & Format$(Val(ATAno), "0000")
          FechaFinal = "30/06/" & Format$(Val(ATAno), "0000")
     Case "2do_Semestre"
          NumMeses = 12
          FechaInicial = "01/07/" & Format$(Val(ATAno), "0000")
          FechaMitad = "15/01/" & Format$(Val(ATAno), "0000")
          FechaFinal = "31/12/" & Format$(Val(ATAno), "0000")
     Case "Todos"
          FechaInicial = "01/01/" & Format$(Val(ATAno), "0000")
          FechaMitad = "15/01/" & Format$(Val(ATAno), "0000")
          FechaFinal = FechaSistema
     Case Else
          NumMeses = LetrasMeses(ATMes)
          FechaInicial = "01/" & Format$(NumMeses, "00") & "/" & Format$(Val(ATAno), "0000")
          FechaMitad = "15/" & Format$(NumMeses, "00") & "/" & Format$(Val(ATAno), "0000")
          FechaFinal = UltimoDiaMes(FechaInicial)
   End Select
   
  'Convertir fecha segun la plataforma
   FechaIni = BuscarFecha(FechaInicial)
   FechaMid = BuscarFecha(FechaMitad)
   FechaFin = BuscarFecha(FechaFinal)
End Sub

Public Function Leer_Archivo_Texto(RutaFile As String) As String
Dim NumFile As Long
Dim TextFile As String
Dim LineFile As String
Dim Results As String
  RatonReloj
  TextFile = ""
  If Len(RutaFile) > 1 Then
     Results = Dir$(RutaFile)
     If Results <> "" Then
        NumFile = FreeFile
        Open RutaFile For Input As #NumFile
        Do While Not EOF(NumFile)
           Line Input #NumFile, LineFile
           TextFile = TextFile & LineFile & " " & vbCrLf
        Loop
        Close #NumFile
     End If
  End If
  RatonNormal
  Leer_Archivo_Texto = Trim(TextFile)
End Function

Public Function Leer_Archivo_Plano(RutaFile As String) As String
Dim NumFile As Long
Dim TextFile As String
Dim LineFile As String
Dim Results As String
  RatonReloj
  TextFile = ""
  If Len(RutaFile) > 1 Then
     Results = Dir$(RutaFile)
     If Results <> "" Then
        NumFile = FreeFile
        Open RutaFile For Input As #NumFile
        Do While Not EOF(NumFile)
           Line Input #NumFile, LineFile
           TextFile = TextFile & LineFile & vbCrLf
        Loop
        Close #NumFile
     End If
  End If
  RatonNormal
  Leer_Archivo_Plano = TextFile
End Function

Public Sub Escribir_Archivo(RutaFile As String, TextFile As String)
Dim NumFile As Long
  RatonReloj
  If Len(RutaFile) > 1 Then
     NumFile = FreeFile
     Open RutaFile For Output As #NumFile
     Print #NumFile, TextFile
     Close #NumFile
  End If
  RatonNormal
End Sub

Public Function Leer_Clave(Login_Gif As String) As String
Dim Clave_Login As String
  Clave_Login = ""
  If Len(Login_Gif) > 1 Then
     I = 1
     Do While I <= Len(Login_Gif)
        Clave_Login = Clave_Login & MidStrg(Login_Gif, I, 1)
        If MidStrg(Login_Gif, I, 1) = "^" Then Clave_Login = ""
        I = I + 1
     Loop
  End If
  Leer_Clave = TrimStrg(Clave_Login)
End Function

Public Sub Abrir_Caja_Registradora()
Open "com1:" For Output As #1 Len = 1
Write #1, Chr(13)
Close #1
End Sub

Public Function ObtenerNombreArchivo(RutaFile As String) As String
Dim CadenaAux As String
   CadenaAux = ""
   J = 0
   If Len(RutaFile) > 3 Then
      For I = Len(RutaFile) To 1 Step -1
          If MidStrg(RutaFile, I, 1) = "\" Then
             J = I + 1: I = 0
          End If
      Next I
      CadenaAux = MidStrg(RutaFile, J, Len(RutaFile))
   End If
   ObtenerNombreArchivo = CadenaAux
End Function

Public Function Obtener_File_Grafico(ArchivoGrafico As String) As String
Dim Resultado As String
    Resultado = ""
    If Len(ArchivoGrafico) > 1 Then
       RutaOrigen = RutaSistema & "\LOGOS\"
       If Len(TrimStrg(Dir$(RutaOrigen & ArchivoGrafico & ".gif"))) > 1 Then
          Resultado = RutaOrigen & ArchivoGrafico & ".gif"
       ElseIf Len(TrimStrg(Dir$(RutaOrigen & ArchivoGrafico & ".jpg"))) > 1 Then
          Resultado = RutaOrigen & ArchivoGrafico & ".jpg"
       ElseIf Len(TrimStrg(Dir$(RutaOrigen & ArchivoGrafico & ".png"))) > 1 Then
          Resultado = RutaOrigen & ArchivoGrafico & ".png"
       End If
    End If
    Obtener_File_Grafico = Resultado
End Function

Public Sub Eliminar_Si_Existe_File(RutaArchivo As String)
'MsgBox "Ruta: " & RutaBuscar & FileBuscar
''    Progreso_Barra.Mensaje_Box = "Eliminando: " & RutaArchivo
''    Progreso_Esperar True
    If Len(TrimStrg(Dir$(RutaArchivo))) Then Kill RutaArchivo
End Sub

Public Function Existe_File(RutaArchivo As String) As Boolean
   'MsgBox "Ruta: " & RutaArchivo
    If Len(RutaArchivo) > 1 Then
       If Len(Dir$(TrimStrg(RutaArchivo))) Then
          Existe_File = True
       Else
          Existe_File = False
       End If
    Else
       Existe_File = False
    End If
End Function

Public Function Existe_Carpeta(RutaCarpeta As String) As Boolean
Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(RutaCarpeta) Then
       Existe_Carpeta = True
    Else
       Existe_Carpeta = False
    End If
    Set fs = Nothing
End Function

Public Function Rubro_Rol_Pago(Detalle_Rol As String) As String
Dim Det_Rol As String
Dim Det_Mitad As String
Dim cod(4) As String
Dim IdCod As Byte
Dim PosCad As Long
 
    Det_Rol = Replace(Detalle_Rol, ".", "")
    Det_Rol = Sin_Signos_Especiales(Det_Rol)
    
    cod(0) = SinEspaciosIzq(Det_Rol)
    Det_Rol = TrimStrg(MidStrg(Det_Rol, Len(cod(0)) + 1, Len(Det_Rol)))
    cod(1) = SinEspaciosIzq(Det_Rol)
    
    Det_Rol = TrimStrg(MidStrg(Det_Rol, Len(cod(1)) + 1, Len(Det_Rol)))
    cod(2) = SinEspaciosIzq(Det_Rol)
    
    Det_Rol = TrimStrg(MidStrg(Det_Rol, Len(cod(2)) + 1, Len(Det_Rol)))
    cod(3) = SinEspaciosIzq(Det_Rol)
    
    Det_Rol = ""
    If Len(TrimStrg(cod(0))) >= 2 Then Det_Rol = Det_Rol & TrimStrg(MidStrg(cod(0), 1, 3)) & "_"
    For IdCod = 1 To 3
        If Len(TrimStrg(cod(IdCod))) >= 2 Then Det_Rol = Det_Rol & TrimStrg(MidStrg(cod(IdCod), 1, 2)) & "_"
    Next
    Det_Rol = TrimStrg(MidStrg(Det_Rol, 1, Len(Det_Rol) - 1))
    If Len(Det_Rol) < 12 Then
       Det_Rol = Replace(Detalle_Rol, ".", "")
       Det_Rol = Sin_Signos_Especiales(Det_Rol)
       Det_Rol = TrimStrg(Det_Rol)
       Select Case Espacios_Blancos(Det_Rol)
         Case 0: Det_Rol = MidStrg(Det_Rol, 1, 12)
         Case 1: Det_Rol = MidStrg(SinEspaciosIzq(Det_Rol), 1, 5) & "_" & MidStrg(SinEspaciosDer(Det_Rol), 1, 5)
         Case 2: PosCad = Len(SinEspaciosIzq(Det_Rol)) + 2
                 Det_Mitad = MidStrg(SinEspaciosIzq(MidStrg(Det_Rol, PosCad, Len(Det_Rol))), 1, 3)
                 Det_Rol = MidStrg(SinEspaciosIzq(Det_Rol), 1, 4) & "_" & Det_Mitad & "_" & MidStrg(SinEspaciosDer(Det_Rol), 1, 3)
         Case Else: Det_Rol = MidStrg(Replace(Det_Rol, " ", "_"), 1, 12)
       End Select
    End If
    Rubro_Rol_Pago = Det_Rol
End Function

Public Function Buscar_Cadena(origen As String, Buscar As String) As Boolean
Dim OrigenTemp As String
    OrigenTemp = Replace(origen, Buscar, "")
    If OrigenTemp <> origen Then Buscar_Cadena = True Else Buscar_Cadena = False
End Function

Public Function Archivo_En_Uso(ByVal sFileName As String) As Boolean
Dim filenum As Integer, errnum As Integer

On Error Resume Next ' Turn error checking off.
filenum = FreeFile() ' Get a free file number.
' Attempt to open the file and lock it.
Open sFileName For Input Lock Read As #filenum
Close filenum ' Close the file.
errnum = Err ' Save the error number that occurred.
On Error GoTo 0 ' Turn error checking back on.

' Check to see which error occurred.
Select Case errnum

' No error occurred.
' File is NOT already open by another user.
Case 0
Archivo_En_Uso = False

' Error number for «Permission Denied.»
' File is already opened by another user.
Case 70
Archivo_En_Uso = True

' Another error occurred.
Case Else
Error errnum
End Select

End Function

Public Function Abreviatura_Texto(Texto As String) As String
Dim NUser1 As String
Dim NUser2 As String
Dim NUser3 As String
Dim NUser4 As String
Dim NUsuario As String
    NUsuario = ""
    If Texto <> Ninguno Then
         NUsuario = Texto
        'Abreviamos el nombre del usuario
         NUser1 = TrimStrg(SinEspaciosIzq(NUsuario))
         NUsuario = TrimStrg(MidStrg(NUsuario, Len(NUser1) + 1, Len(NUsuario)))
         
         NUser2 = TrimStrg(SinEspaciosIzq(NUsuario))
         NUsuario = TrimStrg(MidStrg(NUsuario, Len(NUser2) + 1, Len(NUsuario)))
         
         NUser3 = TrimStrg(SinEspaciosIzq(NUsuario))
         NUsuario = TrimStrg(MidStrg(NUsuario, Len(NUser3) + 1, Len(NUsuario)))
         
         NUser4 = TrimStrg(SinEspaciosIzq(NUsuario))
         NUsuario = TrimStrg(MidStrg(NUsuario, Len(NUser4) + 1, Len(NUsuario)))
         
         NUsuario = MidStrg(NUser1, 1, 1) & "."
         If Len(NUser2) >= 1 Then NUsuario = NUsuario & MidStrg(NUser2, 1, 1) & "."
         If Len(NUser3) >= 1 Then NUsuario = NUsuario & MidStrg(NUser3, 1, 1) & "."
         If Len(NUser4) >= 1 Then NUsuario = NUsuario & MidStrg(NUser4, 1, 1) & "."
         NUsuario = UCaseStrg(TrimStrg(NUsuario))
    End If
    If NUsuario = "" Then NUsuario = Ninguno
    Abreviatura_Texto = NUsuario
End Function

Public Sub PrinterFontBold(Estado As Boolean)
  Printer.FontBold = Estado
End Sub

Public Sub PrinterFontSize(Porte As Integer)
  Printer.FontSize = Porte
End Sub

Public Sub PrinterEndDoc()
  Printer.EndDoc
End Sub

' Función para imprimir texto justificado en impresora.
Public Function Printer_Texto_Justifica(X0, Xf, Y0, txt) As Single
' x0, xf = posicion de los margenes izquierdo y derecho
' y0 = posicion vertical donde se desea empezar a escribir
' txt = texto a escribir
Dim X As Long
Dim y As Long
Dim K As Long
Dim Ancho As Long
Dim S As String, ss As String
Dim x_spc
  S = txt
  X = X0
  y = Y0
  Ancho = (Xf - X0)
  While S <> ""
    ss = ""
    While (S <> "") And (Printer.TextWidth(ss) <= Ancho)
      ss = ss & LeftStrg(S, 1)
      S = RightStrg(S, Len(S) - 1)
    Wend
    If (Printer.TextWidth(ss) > Ancho) Then
       S = RightStrg(ss, 1) & S
       ss = LeftStrg(ss, Len(ss) - 1)
    End If  ' aqui tenemos en ss lo maximo que cabe en una linea
    If RightStrg(ss, 1) = " " Then
       ss = LeftStrg(ss, Len(ss) - 1)
    Else
       If (InStr(ss, " ") > 0) And (LeftStrg(S & " ", 1) <> " ") Then
          While RightStrg(ss, 1) <> " "
            S = RightStrg(ss, 1) & S
            ss = LeftStrg(ss, Len(ss) - 1)
          Wend
          ss = LeftStrg(ss, Len(ss) - 1)
       End If
    End If
    x_spc = 0
    X = X0
    If (Len(ss) > 1) And (S & "" <> "") Then
       x_spc = (Ancho - Printer.TextWidth(ss)) / (Len(ss) - 1)
    End If
    Printer.CurrentX = X
    Printer.CurrentY = y
    If x_spc = 0 Then
       Printer.Print ss;
    Else
       For K = 1 To Len(ss)
         Printer.CurrentX = X
         Printer.Print MidStrg(ss, K, 1);
         X = X + Printer.TextWidth("*" & MidStrg(ss, K, 1) & "*") - Printer.TextWidth("**")
         X = X + x_spc
       Next
    End If
    y = y + Printer.TextHeight(ss)
    While LeftStrg(S, 1) = " "
      S = RightStrg(S, Len(S) - 1)
    Wend
  Wend
  Printer_Texto_Justifica = y + Printer.TextHeight(ss)
End Function
  
'Función para imprimir texto justificado en un PictureBox.
Public Function PictPrint_Texto_Justifica(PictPrint As Object, X0, Xf, Y0, txt) As Single
  ' Muestra un texto justificado dentro del picture "p"
  ' x0, xf = posicion de los margenes izquierdo y derecho
  ' y0 = posicion vertical donde se desea empezar a escribir
  ' txt = texto a escribir
Dim X As Long
Dim y As Long
Dim K As Long
Dim Ancho As Long
Dim AnchoLinea As Long
Dim S As String, ss As String
Dim x_spc
  X0 = X0 + 0.01
  Xf = Xf + 0.01
  S = txt
  X = X0
  y = Y0
  Ancho = (Xf - X0)
 'MsgBox "[" & txt & "]"
  While S <> ""
    ss = ""
    If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
       AnchoLinea = PictPrint.PDFGetTextWidth(ss)
    Else
       AnchoLinea = PictPrint.TextWidth(ss)
    End If
    While (S <> "") And (AnchoLinea <= Ancho)
      ss = ss & LeftStrg(S, 1)
      S = RightStrg(S, Len(S) - 1)
      If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
         AnchoLinea = PictPrint.PDFGetTextWidth(ss)
      Else
         AnchoLinea = PictPrint.TextWidth(ss)
      End If
      'If ss = "" Then ss = " "
    Wend
    If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
       If (PictPrint.PDFGetTextWidth(ss) > Ancho) Then
          S = RightStrg(ss, 1) & S
          ss = LeftStrg(ss, Len(ss) - 1)
       End If
    Else
       If (PictPrint.TextWidth(ss) > Ancho) Then
          S = RightStrg(ss, 1) & S
          ss = LeftStrg(ss, Len(ss) - 1)
       End If
    End If
   'aqui tenemos en ss lo maximo que cabe en una linea
    If RightStrg(ss, 1) = " " Then
       ss = LeftStrg(ss, Len(ss) - 1)
    Else
       If (InStr(ss, " ") > 0) And (LeftStrg(S & " ", 1) <> " ") Then
          While RightStrg(ss, 1) <> " "
            S = RightStrg(ss, 1) & S
            ss = LeftStrg(ss, Len(ss) - 1)
          Wend
          ss = LeftStrg(ss, Len(ss) - 1)
       End If
    End If
    x_spc = 0
    X = X0
    If (Len(ss) > 1) And (S & "" <> "") Then
       If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
          x_spc = (Ancho - PictPrint.PDFGetTextWidth(ss)) / (Len(ss) - 1)
       Else
          x_spc = (Ancho - PictPrint.TextWidth(ss)) / (Len(ss) - 1)
       End If
    End If
    If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
       'PictPrint.PDFTextOut MidStrg(ss, K, 1), X, Y
    Else
       PictPrint.CurrentX = X
       PictPrint.CurrentY = y
    End If
    For K = 1 To Len(ss)
        If TypeOf PictPrint Is mjwPDF Then                'Si la Impresion es en PDF

        Else
           PictPrint.CurrentX = X
        End If
        If MidStrg(ss, K, 1) = "^" Then
           PictPrint.FontBold = True
        ElseIf MidStrg(ss, K, 1) = "~" Then
           PictPrint.FontBold = False
        Else
        If TypeOf PictPrint Is mjwPDF Then                'Si la Impresion es en PDF
           PictPrint.PDFTextOut MidStrg(ss, K, 1), CDbl(X), CDbl(y)
        Else
           PictPrint.Print MidStrg(ss, K, 1);
        End If
           If TypeOf PictPrint Is mjwPDF Then    ' Si la Impresion es en PDF
              X = X + PictPrint.PDFGetTextWidth("*" & MidStrg(ss, K, 1) & "*") - PictPrint.PDFGetTextWidth("**")
           Else
              X = X + PictPrint.TextWidth("*" & MidStrg(ss, K, 1) & "*") - PictPrint.TextWidth("**")
           End If
           X = X + x_spc
        End If
    Next
    If TypeOf PictPrint Is mjwPDF Then                    'Si la Impresion es en PDF
       y = y + PictPrint.PDFTextHeight(ss)
    Else
       y = y + PictPrint.TextHeight(ss)
    End If

    While LeftStrg(S, 1) = " "
      S = RightStrg(S, Len(S) - 1)
    Wend
  Wend

  If TypeOf PictPrint Is mjwPDF Then                  'Si la Impresion es en PDF
     If (y - PictPrint.PDFTextHeight("**")) < Y0 Then
        PictPrint_Texto_Justifica = Y0
     Else
        PictPrint_Texto_Justifica = y - PictPrint.PDFTextHeight("**")
     End If
  Else
     If (y - PictPrint.TextHeight("**")) < Y0 Then
        PictPrint_Texto_Justifica = Y0
     Else
        PictPrint_Texto_Justifica = y - PictPrint.TextHeight("**")
     End If
  End If
End Function

Public Function Leer_Archivo_Textos_Dir(Directorio As String) As Long
Dim ContFile As Long
Dim PosCar As Integer
Dim Archivos As String
  
  ContFile = 0
  Archivos = Dir(Directorio, vbNormal) 'Recupera la primera entrada.
  Do While Archivos <> ""
     If Archivos <> "." And Archivos <> ".." Then ContFile = ContFile + 1
     Archivos = Dir
  Loop
  ReDim Lista_Archivos(ContFile) As String
  ContFile = 0
  Archivos = Dir(Directorio, vbNormal)  'Recupera la primera entrada.
  Do While Archivos <> ""
     If Archivos <> "." And Archivos <> ".." Then
        If (GetAttr(Directorio & Archivos) And vbNormal) = vbNormal Then
           Lista_Archivos(ContFile) = Archivos
        End If
     End If
     Archivos = Dir
     ContFile = ContFile + 1
  Loop
  Leer_Archivo_Textos_Dir = ContFile
End Function

Public Function Listar_Meses()
Listar_Meses = "SELECT * " _
             & "FROM Tabla_Dias_Meses " _
             & "WHERE No_D_M <> 0 " _
             & "AND Tipo = 'M' " _
             & "ORDER BY No_D_M "
End Function

Public Function Listar_Dias()
Listar_Dias = "SELECT * " _
            & "FROM Tabla_Dias_Meses " _
            & "WHERE No_D_M <> 0 " _
            & "AND Tipo = 'D' " _
            & "ORDER BY No_D_M "
End Function

Public Function EsImpar(Numero As Long) As Boolean
   EsImpar = Numero And 1
End Function

Public Sub Exportar_Datagrid_Execel(PathExcel As String)
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
   'Start a new workbook in Excel
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   
   'Add data to cells of the first worksheet in the new workbook
   Set oSheet = oBook.Worksheets(1)
   oSheet.Range("A1").value = "Last Name"
   oSheet.Range("B1").value = "First Name"
   oSheet.Range("A1:B1").Font.bold = True
   oSheet.Range("A2").value = "Doe"
   oSheet.Range("B2").value = "John"
   'Save the Workbook and Quit Excel
   oBook.SaveAs PathExcel
   oExcel.Quit
End Sub

Public Function Leer_Periodo_Lectivo(Optional Campo As String) As Variant
Dim DatosEducativosDB As ADODB.Recordset
Dim Strgs As String
Dim CampoDeTabla As Variant
Dim ContEquiv As Byte

    RatonReloj
    CampoDeTabla = ""
    Rector = Ninguno
    Director = Ninguno
    Secretario1 = Ninguno
    Secretario2 = Ninguno
    Anio_Lectivo = Ninguno
    FormatoLibreta = Ninguno
    Strgs = "SELECT * " _
          & "FROM Catalogo_Periodo_Lectivo " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
    Select_AdoDB DatosEducativosDB, Strgs
    With DatosEducativosDB
     If .RecordCount > 0 Then
         Q_PX = .fields("Q_PP")
         If Q_PX <= 0 Then Q_PX = 80
         Q_PX = Q_PX / 100
         
         Q_EX = .fields("Q_PE")
         If Q_EX <= 0 Then Q_EX = 20
         Q_EX = Q_EX / 100
         
         Cierre_Periodo = .fields("Cierre_Periodo")
         Nota_Mayor = .fields("Nota_Mayor")
         Horas_Min = .fields("Horas_Min")
         Asistencias = .fields("Asistencias")
         Rector = .fields("Rector")
         Director = .fields("Director")
         Secretario1 = .fields("Secretario1")
         Secretario2 = .fields("Secretario2")
         Secretario3 = .fields("Secretario3")
         
         SexoRector = .fields("Sexo_Rector")
         SexoDirector = .fields("Sexo_Director")
         SexoSecre1 = .fields("Sexo_Secre1")
         SexoSecre2 = .fields("Sexo_Secre2")
         SexoSecre3 = .fields("Sexo_Secre3")
         
         TextoSecretario1 = .fields("Texto_Secretario1")
         TextoSecretario2 = .fields("Texto_Secretario2")
         
         TextoBachiller1 = .fields("Bachiller1")
         TextoBachiller2 = .fields("Bachiller2")
         TextoVicerrector1 = .fields("Vicerrector1")
         TextoVicerrector2 = .fields("Vicerrector2")
         TextoRector = .fields("Texto_Rector")
         TextoDirector = .fields("Texto_Director")
         Anio_Lectivo = .fields("Anio_Lectivo")
         FormatoLibreta = .fields("Formato")
         Recomen = .fields("Recomendacion")
         Escalas = .fields("Escala")
         
         Codigo_Ministerio = .fields("Codigo_Colegio")
         Codigo_AMIE = .fields("Codigo_AMIE")
         Mail_Colegio = .fields("Mail_Colegio")
         Institucion1 = .fields("Institucion1")
         Institucion2 = .fields("Institucion2")
         Rubro_Matricula = .fields("Rubro_Matricula")
         Encabezado_Prim = .fields("Encabezado_Prim")
         Encabezado_Secu = .fields("Encabezado_Secu")
         Encabezado_Bach = .fields("Encabezado_Bach")
         TextoWeb = .fields("Web")
         TextoLeyenda = .fields("Leyenda")
         Dec_Nota = .fields("Dec_Nota")
         Tot_Dec_Nota = .fields("Tot_Dec_Nota")
         Suma_Supletorio = .fields("Suma_Supletorio")
         Nota_Rojo = .fields("Nota_Rojo")
         Mejor_Promedio = .fields("Mejor_Promedio")
         Alfabetico = .fields("Alfabetico")
         Distrito = .fields("Distrito")
         Zona = .fields("Zona")
        'Lista de Correos por Secciones
         If Len(.fields("Email_Basica")) > 1 And Len(.fields("Email_Pwd_Basica")) > 1 Then
            Lista_De_Correos(1).Correo_Electronico = .fields("Email_Basica")
            Lista_De_Correos(1).Clave = .fields("Email_Pwd_Basica")
         End If
         If Len(.fields("Email_Secundaria")) > 1 And Len(.fields("Email_Pwd_Secundaria")) > 1 Then
            Lista_De_Correos(2).Correo_Electronico = .fields("Email_Secundaria")
            Lista_De_Correos(2).Clave = .fields("Email_Pwd_Secundaria")
         End If
         If Len(.fields("Email_Bachillerato")) > 1 And Len(.fields("Email_Pwd_Bachillerato")) > 1 Then
            Lista_De_Correos(3).Correo_Electronico = .fields("Email_Bachillerato")
            Lista_De_Correos(3).Clave = .fields("Email_Pwd_Bachillerato")
         End If
                  
        'Logotipo Adicional
         LogoTipo1 = ""
         If .fields("Logo_Tipo") <> Ninguno Then
             RutaOrigen = RutaSistema & "\LOGOS\"
             If Existe_File(RutaOrigen & .fields("Logo_Tipo") & ".gif") Then
                LogoTipo1 = RutaSistema & "\LOGOS\" & .fields("Logo_Tipo") & ".gif"
             Else
                If Existe_File(RutaOrigen & .fields("Logo_Tipo") & ".jpg") Then
                   LogoTipo1 = RutaSistema & "\LOGOS\" & .fields("Logo_Tipo") & ".jpg"
                End If
             End If
         End If
       If Campo <> "" Then CampoDeTabla = .fields(Campo)
     End If
    End With
    DatosEducativosDB.Close
    
   'Cadena de Consulta
    Strgs = "SELECT * " _
          & "FROM Catalogo_Equivalencia " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY Desde, Hasta "
    Select_AdoDB DatosEducativosDB, Strgs
    With DatosEducativosDB
     If .RecordCount > 0 Then
         ContEquiv = 0
         ReDim Equivalencias(.RecordCount) As Tipo_Equivalencias
         Do While Not .EOF
            Equivalencias(ContEquiv).Desde = .fields("Desde")
            Equivalencias(ContEquiv).Hasta = .fields("Hasta")
            Equivalencias(ContEquiv).Rango = .fields("Rango")
            Equivalencias(ContEquiv).Letras = .fields("Letras")
            Equivalencias(ContEquiv).Cualitativa = .fields("Cualitativa")
            Equivalencias(ContEquiv).Cualitativa2 = .fields("Cualitativa2")
            Equivalencias(ContEquiv).Equivalencia = .fields("Equivalencia")
            Equivalencias(ContEquiv).Significado_Letras = .fields("Significado_Letras")
            Equivalencias(ContEquiv).Significado_Letras2 = .fields("Significado_Letras2")
            Equivalencias(ContEquiv).Significado_Evaluacion = .fields("Significado_Evaluacion")
            Equivalencias(ContEquiv).Significado_Evaluacion2 = .fields("Significado_Evaluacion2")
            Equivalencias(ContEquiv).Significado_Equivalencia = .fields("Significado_Equivalencia")
            ContEquiv = ContEquiv + 1
           .MoveNext
         Loop
     Else
         ReDim Equivalencias(1) As Tipo_Equivalencias
         Equivalencias(0).Desde = 0
         Equivalencias(0).Hasta = 0
         Equivalencias(0).Letras = Ninguno
         Equivalencias(0).Cualitativa = Ninguno
         Equivalencias(0).Rango = Ninguno
         Equivalencias(0).Equivalencia = Ninguno
         Equivalencias(0).Significado_Letras = Ninguno
         Equivalencias(0).Significado_Evaluacion = Ninguno
         Equivalencias(0).Significado_Equivalencia = Ninguno
     End If
    End With
    DatosEducativosDB.Close
    RatonNormal
    Leer_Periodo_Lectivo = CampoDeTabla
End Function

Public Function Sin_Signos_Especiales(Cadena As String) As String
Dim CadResult As String
Dim CadTemp As String
Dim utf8Bytes() As Byte
Dim Idc As Long
    CadResult = ""
    
    If Len(Cadena) > 0 Then
      'MsgBox "Destktop Test>: " & Cadena
       CadTemp = Cadena
       CadTemp = Replace(CadTemp, "á", "a")
       CadTemp = Replace(CadTemp, "é", "e")
       CadTemp = Replace(CadTemp, "í", "i")
       CadTemp = Replace(CadTemp, "ó", "o")
       CadTemp = Replace(CadTemp, "ú", "u")
       CadTemp = Replace(CadTemp, "ñ", "n")
       CadTemp = Replace(CadTemp, "Ñ", "N")
       CadTemp = Replace(CadTemp, "Á", "A")
       CadTemp = Replace(CadTemp, "Í", "I")
       CadTemp = Replace(CadTemp, "Ó", "O")
       CadTemp = Replace(CadTemp, "Ú", "U")

       CadTemp = Replace(CadTemp, "Nº", "No.")
       CadTemp = Replace(CadTemp, "ª", "a. ")
       CadTemp = Replace(CadTemp, "°", "o. ")
       CadTemp = Replace(CadTemp, "½", "1/2")
       CadTemp = Replace(CadTemp, "½", "1/4")
       CadTemp = Replace(CadTemp, "´", "`")
       CadTemp = Replace(CadTemp, "&", " Y ")
    
       CadTemp = TrimStrg(CadTemp)
       CadTemp = Replace(CadTemp, "  ", " ")
       CadTemp = Replace(CadTemp, "   ", " ")
       CadTemp = Replace(CadTemp, vbCr, "|")
       CadTemp = Replace(CadTemp, vbLf, "|")

''      'Convertir la cadena a bytes usando la codificación Unicode (UTF-16)
''       utf8Bytes = StrConv(CadTemp, vbFromUnicode)
''       For Idc = 0 To UBound(utf8Bytes)
''           MsgBox Cadena & " => " & utf8Bytes(Idc) & "=" & Chr(utf8Bytes(Idc))
''           Select Case utf8Bytes(Idc)
''             Case 9, 10, 11, 13: CadResult = CadResult & Chr(utf8Bytes(Idc))
''             Case 32 To 126: CadResult = CadResult & Chr(utf8Bytes(Idc))
''             Case 181: CadResult = CadResult & Chr(utf8Bytes(Idc)) ' Á
''             Case 144: CadResult = CadResult & Chr(utf8Bytes(Idc)) ' É
''             Case 214: CadResult = CadResult & Chr(utf8Bytes(Idc)) ' Í
''             Case 224: CadResult = CadResult & Chr(utf8Bytes(Idc)) ' Ó
''             Case 233: CadResult = CadResult & Chr(utf8Bytes(Idc)) ' Ú
''             Case Else: CadResult = CadResult & " "
''           End Select
''       Next Idc
''    CadResult = TrimStrg(CadResult)
''    CadResult = Replace(CadResult, "  ", " ")
''    CadResult = Replace(CadResult, "   ", " ")
''    CadResult = Replace(CadResult, vbCr, "|")
''    CadResult = Replace(CadResult, vbLf, "|")
    End If
    CadResult = CadTemp
'    Clipboard.Clear
'    Clipboard.SetText CadResult
    ' MsgBox "Destktop Test<: " & CadResult
    Sin_Signos_Especiales = CadResult
End Function

Public Function Sin_Signos_XML(Cadena As String) As String
Dim Cad As String
'ParaAux = Replace(ParaAux, "|", "")
    Cad = TrimStrg(Cadena)
    Cad = Replace(Cad, "á", "a")
    Cad = Replace(Cad, "é", "e")
    Cad = Replace(Cad, "í", "i")
    Cad = Replace(Cad, "ó", "o")
    Cad = Replace(Cad, "ú", "u")
    Cad = Replace(Cad, "Á", "A")
    Cad = Replace(Cad, "É", "E")
    Cad = Replace(Cad, "Í", "I")
    Cad = Replace(Cad, "Ó", "O")
    Cad = Replace(Cad, "Ú", "U")
    Cad = Replace(Cad, "à", "a")
    Cad = Replace(Cad, "è", "e")
    Cad = Replace(Cad, "ì", "i")
    Cad = Replace(Cad, "ò", "o")
    Cad = Replace(Cad, "ù", "u")
    Cad = Replace(Cad, "À", "A")
    Cad = Replace(Cad, "È", "E")
    Cad = Replace(Cad, "Ì", "I")
    Cad = Replace(Cad, "Ò", "O")
    Cad = Replace(Cad, "Ù", "U")
    Cad = Replace(Cad, "ñ", "n")
    Cad = Replace(Cad, "Ñ", "N")
    Cad = Replace(Cad, "ü", "u")
    Cad = Replace(Cad, "Ü", "U")
    Cad = Replace(Cad, "&", "Y")
    Cad = Replace(Cad, "Nº", "No.")
    Cad = Replace(Cad, "ª", "a. ")
    Cad = Replace(Cad, "°", "o. ")
    Cad = Replace(Cad, "½", "1/2")
    Cad = Replace(Cad, "½", "1/4")
    Cad = Replace(Cad, Chr(255), " ")
    Cad = Replace(Cad, Chr(254), " ")
    Cad = Replace(Cad, "^", "")
    Sin_Signos_XML = Cad
End Function

'FoxitCtl
''      VerPDF.OpenFile " "
''      VerPDF.OpenFile RutaSysBases & "\TEMP\" & tPrint.NombreArchivo & ".pdf"
'''Public Sub Presentar_PDF(VerPDF As AcroPDF)
'''   If SetNombrePRN = Impresota_PDF And tPrint.VerDocumento = False Then
'''      RatonReloj
'''      'MsgBox tPrint.NombreArchivo
'''      With VerPDF
'''          .setZoom 125
'''          .setShowScrollbars True
'''          .setShowToolbar False
'''          .LoadFile RutaSysBases & "\TEMP\" & tPrint.NombreArchivo & ".pdf"
'''      End With
'''      RatonNormal
'''   End If
'''End Sub

'Tiene que venir como parametros en la variable RutaDocumentoPDF el archivo
'''Public Sub Presentar_PDF(VerPDF As AcroPDF, RutaArchivoPDF As String, Optional PorcZoom As Byte)
'''    RatonReloj
'''    'If PorcZoom = 0 Then VerPDF.setZoom 125 Else VerPDF.setZoom PorcZoom
'''    'VerPDF.setShowScrollbars True
'''    'VerPDF.setShowToolbar False
'''
'''    If RutaArchivoPDF = "" Then
'''      'MsgBox "No existe archivo que presentar," & vbCrLf & "revise que de verdad existe."
'''       VerPDF.LoadFile "C:\"
'''    Else
'''       If Existe_File(RutaArchivoPDF) Then
'''          VerPDF.LoadFile RutaArchivoPDF
'''       Else
'''          RatonNormal
'''          MsgBox "EL ARCHIVO:" & vbCrLf & vbCrLf & RutaArchivoPDF & vbCrLf & vbCrLf & "NO EXISTE O ESTA DAÑADO"
'''          VerPDF.LoadFile "C:\"
'''       End If
'''    End If
'''    RatonNormal
'''End Sub

Public Function Extraer_Apellidos(Beneficiario As String) As String
 Dim I As Long
 Dim ContStr As Byte
 Dim Result As String
 If Len(Beneficiario) > 1 Then
    ContStr = 0
    For I = 1 To Len(Beneficiario)
        If MidStrg(Beneficiario, I, 1) = " " Then ContStr = ContStr + 1
        If ContStr = 2 Then
           Result = TrimStrg(MidStrg(Beneficiario, 1, I))
           ContStr = ContStr + 1
        End If
    Next I
 Else
    Result = ""
 End If
 Extraer_Apellidos = Result
End Function

Public Function Extraer_Nombres(Beneficiario As String) As String
 Dim I As Long
 Dim ContStr As Byte
 Dim Result As String
 Result = ""
 If Len(Beneficiario) > 1 Then
    ContStr = 0
    For I = 1 To Len(Beneficiario)
        If MidStrg(Beneficiario, I, 1) = " " Then ContStr = ContStr + 1
        If ContStr = 2 Then
           Result = TrimStrg(MidStrg(Beneficiario, I, Len(Beneficiario)))
           ContStr = ContStr + 1
        End If
    Next I
 End If
 Extraer_Nombres = Result
End Function

'''Public Sub Progreso_Esperar(Optional Mensaje As String)
'''Dim ValorPorc As Single
'''  With FEsperar
'''       If Progreso_Barra.Puntos < 0 Then Progreso_Barra.Puntos = 0
'''       If Progreso_Barra.Puntos > 20 Then Progreso_Barra.Puntos = 0
'''       Progreso_Barra.Puntos = Progreso_Barra.Puntos + 1
'''       If Progreso_Barra.Incremento < 0 Then Progreso_Barra.Incremento = 0
'''       If Progreso_Barra.Valor_Maximo <= 0 Then Progreso_Barra.Valor_Maximo = 1
'''       Progreso_Barra.Incremento = Progreso_Barra.Incremento + 1
'''      .Refresh
'''       If Progreso_Barra.Incremento <= Progreso_Barra.Valor_Maximo Then
'''         .Label2.Caption = String(Progreso_Barra.Puntos, ".") & "Procesando el " _
'''                         & Format(Progreso_Barra.Incremento / Progreso_Barra.Valor_Maximo, "00%") & ", Espere por favor" _
'''                         & String(Progreso_Barra.Puntos, ".")
'''       Else
'''         .Label2.Caption = String(Progreso_Barra.Puntos, ".") & "Procesando el 100%, Espere por favor" & String(Progreso_Barra.Puntos, ".")
'''          Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
'''       End If
'''      .Refresh
'''       If Mensaje <> "" Then
'''          Progreso_Barra.Mensaje_Box = Mensaje
'''       Else
'''          Progreso_Barra.Mensaje_Box = "G e n e r a n d o"
'''       End If
'''      .Label1.Caption = Progreso_Barra.Mensaje_Box
'''      .Refresh
'''  End With
'''End Sub
'''''''''''''''''''''''

Function Campo_Blanco(Dato As String) As String
  If Len(Dato) > 1 And Val(Dato) <> 0 Then Campo_Blanco = Format(Val(Dato), "##0.00") Else Campo_Blanco = ""
End Function

'Web Service
Function InvokeWebService(strSoap, strSOAPAction, strURL, ByRef xmlResponse) As Boolean
Dim xmlhttp As MSXML2.XMLHTTP30
Dim blnSuccess As Boolean

Set xmlhttp = New MSXML2.XMLHTTP30
xmlhttp.open "POST", strURL, False
xmlhttp.setRequestHeader "Man", "POST " & strURL & " HTTP/1.1"
xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
xmlhttp.setRequestHeader "SOAPAction", strSOAPAction
Call xmlhttp.send(strSoap)

If xmlhttp.Status = 200 Then
blnSuccess = True
Else
blnSuccess = False
End If

Set xmlResponse = xmlhttp.responseXML
InvokeWebService = blnSuccess
Set xmlhttp = Nothing
End Function

Public Sub Imagen_Esperar(Optional MensajeEsperar As String)
On Error Resume Next
Dim MyTimer As Integer
Dim I As Long
    If IsFormLoaded(FEsperar) Then
       RatonReloj
''       FrameCount = FrameCount + 1
''       MyTimer = Time
       If Len(MensajeEsperar) > 1 Then
          FEsperar.LblMensaje.Caption = MensajeEsperar & "..."
       Else
          FEsperar.LblMensaje.Caption = "Procesando..."
       End If
       FEsperar.Refresh
       If FrameCount < TotalFrames Then
          FEsperar.Image1(FrameCount).Visible = False
          FrameCount = FrameCount + 1
       Else
          FrameCount = 0
          For I = 1 To FEsperar.Image1.Count - 1
              FEsperar.Image1(I).Visible = False
          Next I
       End If
       FEsperar.Image1(FrameCount).Visible = True
       FEsperar.Timer1.Interval = CLng(FEsperar.Image1(FrameCount).Tag)
       FEsperar.Refresh
    Else
       RatonNormal
    End If
    If Err Then Exit Sub
End Sub

Public Sub TVAddNode(ByRef XML_Node As IXMLDOMNode, _
                     ByRef TVTreeView As TreeView, _
                     Optional ByRef TreeNode As node)
Dim xNode As node
Dim xNodeList As IXMLDOMNodeList
Dim I As Long
    
    If TreeNode Is Nothing Then
       Set xNode = TVTreeView.Nodes.Add
    Else
       Set xNode = TVTreeView.Nodes.Add(TreeNode, tvwChild)
    End If
    
    xNode.Expanded = True
    xNode.Text = XML_Node.nodeName
    
    If xNode.Text = "#text" Then
       xNode.Text = XML_Node.nodeTypedValue
    Else
      ' xNode.Text = "<" + xNode.Text + ">"
       xNode.Text = xNode.Text
    End If
    
    Set xNodeList = XML_Node.childNodes
    For I = 0 To xNodeList.Length - 1
        TVAddNode xNodeList.Item(I), TVTreeView, xNode
    Next
End Sub

Function Espacios_Blancos(Texto As String) As Byte
Dim NumBlanco As Byte
Dim Ic As Long
   NumBlanco = 0
   For Ic = 1 To Len(Texto)
       If MidStrg(Texto, Ic, 1) = " " Then NumBlanco = NumBlanco + 1
   Next Ic
   Espacios_Blancos = NumBlanco
End Function

Public Function Load_Gif(sFile As String, aImg As Variant) As Long
On Error Resume Next

Dim hFile As Long
Dim sImgHeader As String
Dim sFileHeader As String
Dim sBuff As String
Dim sPicsBuff As String
Dim nImgCount As Long
Dim I As Long
Dim J As Long
Dim K As Long
Dim Ki As Long
Dim Kf As Long
Dim xOff As Long
Dim yOff As Long
Dim TimeWait As Long
Dim sGifMagic As String
Dim imgTemp As String

If Dir$(sFile) = "" Or sFile = "" Then
    MsgBox "Fila " & sFile & " no encontrada", vbInformation
    Exit Function
End If

sGifMagic = Chr$(0) & Chr$(33) & Chr$(249)

If aImg.Count > 1 Then
    For I = 1 To aImg.Count - 1
        Unload aImg(I)
    Next I
End If

hFile = FreeFile

Open sFile For Binary Access Read As hFile
    sBuff = String$(LOF(hFile), Chr(0))
    Get #hFile, , sBuff
Close #hFile

I = 1
nImgCount = 0
J = InStr(1, sBuff, sGifMagic) + 1
sFileHeader = LeftStrg(sBuff, J)

If LeftStrg(sFileHeader, 3) <> "GIF" Then
    MsgBox "No es una fila *.gif", vbInformation
    Exit Function
End If

Load_Gif = True

I = J + 2

If Len(sFileHeader) >= 127 Then
    RepeatTimes = Asc(MidStrg(sFileHeader, 126, 1)) + (Asc(MidStrg(sFileHeader, 127, 1)) * 256&)
Else
    RepeatTimes = 0
End If

For K = Len(sFile) To 1 Step -1
    If MidStrg(sFile, K, 1) = "." Then Kf = K
    If MidStrg(sFile, K, 1) = "\" Then
       Ki = K + 1
       Exit For
    End If
Next

hFile = FreeFile
Open "temp.gif" For Binary As hFile
Do
    imgTemp = RutaSysBases & "\TEMP\" & MidStrg(sFile, Ki, Kf - Ki) & Format(nImgCount, "00") & ".gif"
    nImgCount = nImgCount + 1
    J = InStr(I, sBuff, sGifMagic) + 3
    If J > Len(sGifMagic) Then
        sPicsBuff = String$(Len(sFileHeader) + J - I, Chr$(0))
        sPicsBuff = sFileHeader & MidStrg(sBuff, I - 1, J - I)
        Put #hFile, 1, sPicsBuff
        
        sImgHeader = LeftStrg(MidStrg(sBuff, I - 1, J - I), 16)
        TimeWait = ((Asc(MidStrg(sImgHeader, 4, 1))) + (Asc(MidStrg(sImgHeader, 5, 1)) * 256&)) * 10&
        
        If nImgCount > 1 Then
            Load aImg(nImgCount - 1)
            xOff = Asc(MidStrg(sImgHeader, 9, 1)) + (Asc(MidStrg(sImgHeader, 10, 1)) * 256&)
            yOff = Asc(MidStrg(sImgHeader, 11, 1)) + (Asc(MidStrg(sImgHeader, 12, 1)) * 256&)

            aImg(nImgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
            aImg(nImgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
        End If
        aImg(nImgCount - 1).Tag = TimeWait
        aImg(nImgCount - 1).Picture = LoadPicture("temp.gif")
        'SavePicture aImg(nImgCount - 1).Picture, imgTemp
        I = J
    End If
Loop Until J = 3
Close #hFile
'MsgBox nImgCount

Kill "temp.gif"

TotalFrames = aImg.Count - 1

Load_Gif = TotalFrames
Exit Function

errHandler:
MsgBox "Error No. " & Err.Number & " leyendo la fila", vbCritical
Load_Gif = False
On Error GoTo 0
If Err Then Beep
End Function

Public Function consulta_RUC_SRI(NumRUC As String) As Tipo_Contribuyente
Dim Cont As Integer
Dim IniNodo As Integer
Dim FinNodo As Integer
Dim vNodos() As String
Dim ResultURL As String
Dim Result As Tipo_Contribuyente

    RatonReloj
    With Result
        .Existe = False
        .Estado = Ninguno
        .RUC_SRI = Ninguno
        .RazonSocial = Ninguno
        .NombreComercial = Ninguno
        .ClaseRUC = Ninguno
        .TipoRUC = Ninguno
        .Obligado = "SI"
        .ActividadEconomica = Ninguno
        .Categoria = Ninguno
        .FechaInicio = Ninguno
        .FechaCese = Ninguno
        .FechaReinicio = Ninguno
        .FechaActualización = Ninguno
        .AgenteRetencion = Ninguno
        .MicroEmpresa = Ninguno
    End With
    
    If Len(NumRUC) = 13 And Ping_IP("srienlinea.sri.gob.ec") And GetUrlSource(urlEsUnRUC & NumRUC) = "true" Then
      'Verificamos que tipo de contribuyente es
       Tipo_Contribuyente_SP_MySQL NumRUC, Result.MicroEmpresa, Result.AgenteRetencion
       Result.Existe = True
       XML = Replace(GetUrlSource(urlDatosDelRUC & NumRUC), """", "'")
       
      'Generar_File_SQL "SRI_LINEA_2", XML
       XML = Replace(XML, """", "'")
       Cont = InStr(XML, "<table class='formulario'>")
       If Cont > 1 Then
          XML = Mid(XML, Cont, Len(XML))
          Cont = InStr(XML, "</table>")
          XML = Mid(XML, 1, Cont + 8)
          XML = Trim(XML)
           'Escribir_Archivo "c:\temp\index_" & NumRUC & ".php", XML
            Cont = 0
            Do While Len(XML) > 0
               ReDim Preserve vNodos(Cont) As String
               vNodos(Cont) = Ninguno
               IniNodo = InStr(XML, "<td>")
               FinNodo = InStr(XML, "</td>")
               If IniNodo > 0 And FinNodo > 0 Then
                  If IniNodo < FinNodo Then
                     vNodos(Cont) = Mid(XML, IniNodo + 4, FinNodo - IniNodo - 4)
                     vNodos(Cont) = Replace(vNodos(Cont), vbCrLf, "")
                     vNodos(Cont) = Replace(vNodos(Cont), vbCr, "")
                     vNodos(Cont) = Replace(vNodos(Cont), vbLf, "")
                     vNodos(Cont) = Replace(vNodos(Cont), vbTab, "")
                     vNodos(Cont) = Replace(vNodos(Cont), "&nbsp;", "")
                     vNodos(Cont) = Replace(vNodos(Cont), "</a>", "")
                     vNodos(Cont) = Replace(vNodos(Cont), "<a class='link2' href='javascript:sociedad();'", "")
                     vNodos(Cont) = Replace(vNodos(Cont), "onclick='forma.ruc.value='" & NumRUC & "''>", "")
                     vNodos(Cont) = Trim(UCase(vNodos(Cont)))
                  End If
                  XML = Mid(XML, FinNodo + 4, Len(XML))
               Else
                  XML = ""
               End If
               Cont = Cont + 1
            Loop
            With Result
                .RazonSocial = vNodos(0)
                .RUC_SRI = vNodos(1)
                .NombreComercial = vNodos(2)
                .Estado = vNodos(4)
                .ClaseRUC = vNodos(5)
                .TipoRUC = vNodos(6)
                 Cont = 7
                 If Len(vNodos(Cont)) = 2 Then
                   .Obligado = vNodos(Cont)
                    Cont = Cont + 1
                 End If
                .ActividadEconomica = vNodos(Cont)
                 Cont = Cont + 1
                .FechaInicio = vNodos(Cont)
                 Cont = Cont + 1
                .FechaCese = vNodos(Cont)
                 Cont = Cont + 1
                .FechaReinicio = vNodos(Cont)
                 Cont = Cont + 1
                .FechaActualización = vNodos(Cont)
                 Cont = Cont + 1
                .Categoria = vNodos(Cont)
            End With
       End If
    Else
       Result.Estado = "NO ES RUC VALIDO"
       Result.Obligado = Ninguno
    End If
    RatonNormal
    consulta_RUC_SRI = Result
End Function

'''Public Function Estado_Servidor(IPServidor As String) As Boolean
'''Dim Resultado As Boolean
'''    Resultado = True
'''    If IP_PC.InterNet Then
'''       If Not Ping_IP(IPServidor) Then
'''          MsgBox "LA CONEXION NO ESTA ESTABLECIDA, POR FAVOR LLAME AL ADMINISTRADOR DEL SISTEMA PARA QUE ACTIVE EL SERVIDOR " & vbCrLf & IP_PC.Status
'''          Resultado = False
'''       End If
'''    Else
'''       MsgBox "EN ESTOS MOMENTOS NO CUENTA CON CONEXION A INTERNET, INTENTE MAS TARDE QUE SE ACTIVE LA CONEXION A INTERNET" & vbCrLf & IP_PC.Status
'''       Resultado = False
'''    End If
'''    Estado_Servidor = Resultado
'''End Function

Public Function SelectDialogFile(Optional RutaCarpeta As String, Optional FiltroFile As String) As String
   With MDIFormulario.Dir_Dialog
       .Filename = ""
       .Flags = cdlOFNFileMustExist + cdlOFNNoChangeDir + cdlOFNHideReadOnly
       .CancelError = False
       .PrinterDefault = False
        If Len(RutaCarpeta) > 1 Then .InitDir = RutaCarpeta Else .InitDir = RutaSysBases
       .DialogTitle = "Seleccione un Archivo"
        If Len(FiltroFile) > 1 Then .Filter = FiltroFile Else .Filter = "Archivos de texto|*.txt|Archivos Excel|*.xls|Archivos CSV|*.csv"
       .ShowOpen
       If .Filename = "" Then MsgBox "No se ha seleccionado ningún archivo", vbInformation '.Filename = "Seleccione un Archivo"
       SelectDialogFile = .Filename
   End With
End Function

