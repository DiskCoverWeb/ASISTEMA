Attribute VB_Name = "Modulos"
Option Explicit

' Estructura SHFILEOPSTRUCT o para usar con el Api
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

'Constantes
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40

'Constantes para las teclas y otros
Private Const KEYEVENTF_KEYUP = &H2
Private Const KEYEVENTF_EXTENDEDKEY = &H1

'Declaración Api SHFileOperation
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
                                              ByVal bScan As Byte, _
                                              ByVal dwFlags As Long, _
                                              ByVal dwExtraInfo As Long)

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
Global Const TadMemo = adLongVarWChar     ' NText


'-------------------------------------
' Boolean
Global File_Emails As Boolean
Global Evaluar As Boolean
Global SQL_Server As Boolean
Global DiaZip(1 To 7) As Boolean

' Byte
Global PC_Numero As Byte
Global DiasRespaldo As Byte
Global contadorEmail As Byte
Global Cantidad_Cyber_Tiempo As Byte

' Integer
Global NumFile As Integer

' Long
Global Erg As Long
Global NumPos As Long
Global Minutos As Long
Global Segundos As Long
Global Contador As Long
Global FileResp As Long
Global TotalReg As Long
Global ContadorRUCCI As Long
Global MDI_X_Max As Long
Global MDI_Y_Max As Long

' Single
Global MiTiempo As Single
Global IntervaloTiempo As Single

' Fecha o Date
Global TiempoSistema As Date

' String
Global sSQL As String
Global HoraSistema As String
Global LineaTexto As String
Global RutaOrigen As String
Global RutaGeneraFile As String
Global RutaDestino As String
Global RutaSistema As String
Global RutaSysBases As String
Global RutaUpdate As String
Global Email_Respaldo As String
Global Cadena As String
Global ProveedorAccess As String
Global AdoStrCnn As String
Global AdoStrCnnAccess As String
Global RutaEmpresa As String
Global RutaEmpresaOld As String
Global PathEmpresa As String
Global FechaSistema As String
Global NumEmpresa As String
Global CodigoUsuario As String
Global NombreUsuario As String
Global Empresa As String
Global Periodo_Contable As String
Global Unidad As String
Global LineaConexion(5) As String
Global CadAux As String
Global FechaRespaldo As String
Global MiArchivo, MiRuta, MiNombre
Global Modulo As String
Global WidthBoolean As Single
Global WidthText As Single
Global WidthDate As Single
Global WidthTime As Single
Global WidthByte As Single
Global WidthLong As Single
Global WidthInteger As Single
Global WidthCurrency As Single
Global WidthSingle As Single
Global WidthDouble  As Single

'Ancho de Tipos Ado
Global CadBoolean As String
Global CadDate As String
Global CadDate1 As String
Global CadTime As String
Global CadByte As String
Global CadInteger As String
Global CadLong As String
Global CadSingle As String
Global CadDouble As String
Global CadCurrency As String
Global Nombre_Base_SQL As String
Global CarIni As Integer
Global CarFin As Integer
Global RegAfectados As Long
Global Nombre_Base_Respaldo As String

'Tipo de tiempo
Global VCyber_Tiempo() As Cyber_Tiempo
Global Lista_De_Correos(4) As Tipo_Lista_Mail
Global TipoC() As Campos_Tabla

Global Vect_Dec() As Campos_Decimal
Global NombreDeEmpresas() As String
Global RUCDeEmpresas() As String
Global TMail As Tipo_Mail
'-------------------------------------
' Rutinas Generales
Global ShiftDown As Boolean
Global AltDown As Boolean
Global CtrlDown As Boolean

Public Sub RatonReloj()
  Screen.MousePointer = vbHourglass
End Sub

Public Sub RatonNormal()
  Screen.MousePointer = vbDefault
End Sub

Public Function MesesLetras(Mes As Integer) As String
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
   MesesLetras = SMes
End Function

Public Function DiasLetras(Mes As Byte) As String
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

Public Sub SelectAdodc(DataSQL As Adodc, _
                       SQLs As String, _
                       Optional SetMsg As Boolean)
Dim Strgs As String
   If SQLs <> "" Then
      RatonReloj
      SQLs = CompilarSQL(SQLs)
     'MsgBox SQLs
      DataSQL.RecordSource = SQLs
      DataSQL.Refresh
'If SQL_Server = False Then DataSQL.Refresh
      RatonNormal
      With DataSQL.Recordset
      If .RecordCount > 0 Then
         .MoveLast
          Strgs = "CONSULTA PROCESADA CORRECTAMENTE." & vbCrLf & vbCrLf _
                & "Existen " & .RecordCount & " Registro(s) Procesado(s)."
          If SetMsg Then MsgBox Strgs
          Strgs = "Registros: " & Format(.RecordCount, "#,##0") & "." _
                & Space(30) & "Página(s): " _
                & Format((.RecordCount / 45) + 1, "#,##0") & "."
          DataSQL.Caption = Strgs
         .MoveFirst
      Else
         Strgs = "No existen Datos Disponibles." & vbCrLf _
               & "'Si cree tener datos, consulte al Técnico'."
         If SetMsg Then MsgBox Strgs
      End If
      End With
   End If
End Sub

Public Sub Determina_Ancho_Tipos(Optional EsCorto As Boolean)
  If EsCorto Then
     CadBoolean = "Si "
     CadDate = "dd/MM/aaaa "
     CadDate1 = "dd/MM/aaaa "
     CadTime = "hh:mm "
     CadByte = "999 "
     CadInteger = "9999 "
     CadLong = "999999 "
     CadSingle = "999.00% "
     CadDouble = "9,999.00 "
     CadCurrency = "9,999,999.00 "
  Else
     CadBoolean = "Yes "
     CadDate = "dd/mm/yyyy "
     CadDate1 = "dd/mm/yyyy "
     CadTime = "HH:MM:SS "
     CadByte = "+999 "
     CadInteger = "+99999 "
     CadLong = "+99999999 "
     CadSingle = "+999.00% "
     CadDouble = "+99,999,999.00 "
     CadCurrency = "+9,999,999,999.00 "
  End If
End Sub

Public Sub Select_Adodc_Grid(DBGMalla As DataGrid, _
                             Dtas As Adodc, _
                             SQuerys As String, _
                             Optional Decimales As String, _
                             Optional EsCampoCorto As Boolean, _
                             Optional PresentarEsperar As Boolean, _
                             Optional NombreFile As String)
Dim AnchoMax As Byte
Dim Presentar As Boolean
Dim LenCamposDec As Integer
Dim CantDecCampo As String
Dim Otros_Dec As Boolean
Dim NumFile As Long
Dim CantDec As Long
Dim CantCampos As Long
Dim C1 As Long
Dim C2 As Long
Dim Col As Long
Dim width As Single
Dim maxWidth As Single
Dim celdaText As String
Dim saveFont As StdFont
Dim oldScaleMode As Integer
Dim Porcentaje As Single
Dim CadAux As String
Dim DatosFile As String

On Error GoTo Errorhandler
 
   If SQuerys <> "" Then
      RatonReloj
     'Determinamos el ancho de los tipo de campo que tiene VB
      Determina_Ancho_Tipos EsCampoCorto
      WidthBoolean = CSng(DBGMalla.Parent.TextWidth(CadBoolean))
      WidthDate = CSng(DBGMalla.Parent.TextWidth(UCase(CadDate)))
      WidthTime = CSng(DBGMalla.Parent.TextWidth(CadTime))
      WidthByte = CSng(DBGMalla.Parent.TextWidth(CadByte))
      WidthInteger = CSng(DBGMalla.Parent.TextWidth(CadInteger))
      WidthLong = CSng(DBGMalla.Parent.TextWidth(CadLong))
      WidthSingle = CSng(DBGMalla.Parent.TextWidth(CadSingle))
      WidthDouble = CSng(DBGMalla.Parent.TextWidth(CadDouble))
      WidthCurrency = CSng(DBGMalla.Parent.TextWidth(CadCurrency))
      
      Presentar = DBGMalla.Visible
      DBGMalla.Visible = False
      
      SQuerys = CompilarSQL(SQuerys)
     'MsgBox AdoStrCnn
     'MsgBox SQuerys
      Dtas.RecordSource = SQuerys
      Dtas.Refresh

     'Variables para la cantidad de registros y columnas
      CantCampos = Dtas.Recordset.Fields.Count
      
     'MsgBox Dtas.Recordset.Fields.Count
     'Si el número de registros es igual a 0 salimos
     'If Dtas.Recordset.RecordCount = 0 Then Exit Sub
     
     'Guardamos la fuente del DataGrid para luego reestablecerla
      Set saveFont = DBGMalla.Parent.Font
      Set DBGMalla.Parent.Font = DBGMalla.Font
     'Ajustar el ScaleMode en vbTwips para el formulario
      oldScaleMode = DBGMalla.Parent.ScaleMode
''      DBGMalla.Parent.ScaleMode = vbTwips
''      DBGMalla.Refresh
     'Array para almacenar el ancho de cada columna
      ReDim Vect_Dec(CantCampos) As Campos_Decimal
     'Enceramos la impresion
      DBGMalla.HeadFont.Bold = True
      For Col = 0 To CantCampos - 1
          Vect_Dec(Col).Campo = Dtas.Recordset.Fields(Col).Name
          Vect_Dec(Col).CantDec = 2
          Vect_Dec(Col).AnchoCampo = DBGMalla.Parent.TextWidth(UCase(Dtas.Recordset.Fields(Col).Name & " "))
      Next Col
     'Determinados si hay que presentar mas de dos decimales a los numeros
      CadAux = Decimales
      For Col = 0 To CantCampos - 1
         'MsgBox Decimales & vbCrLf & Vect_Dec(Col).Campo & vbCrLf & Vect_Dec(Col).CantDec & vbCrLf & Vect_Dec(Col).AnchoCampo
          If Len(CadAux) > 2 Then
             C1 = InStr(CadAux, Vect_Dec(Col).Campo)
             If C1 > 0 Then
                C1 = InStr(CadAux, " ") + 1
                C2 = InStr(CadAux, "|") - 1
               'MsgBox Vect_Dec(Col).Campo & vbCrLf & MidStrg(CadAux, C1, C2 - C1 + 1)
                Vect_Dec(Col).CantDec = Val(Mid(CadAux, C1, C2 - C1 + 1))
                CadAux = Trim(Mid(CadAux, C2 + 2, Len(CadAux)))
             End If
          End If
         ' MsgBox " -> " & Vect_Dec(Col).Campo & vbCrLf & Vect_Dec(Col).CantDec
      Next Col
      For Col = 0 To CantCampos - 1
          With DBGMalla
               'MsgBox Col & vbCrLf & Vect_Dec(Col).AnchoCampo & vbCrLf & Dtas.Recordset.Fields(Col).Name
               'Vect_Dec(Col).AnchoCampo = DBGMalla.Parent.TextWidth(Dtas.Recordset.Fields(Col).Name)
               Vect_Dec(Col).SumaTotal = 0
               'MsgBox Vect_Dec(Col).AnchoCampo & vbCrLf & Dtas.Recordset.Fields(Col).Name
               Select Case Dtas.Recordset.Fields(Col).Type
                 Case TadBoolean
                     .Columns(Col).NumberFormat = Format$("Yes/No")
                      If Vect_Dec(Col).AnchoCampo < WidthBoolean Then Vect_Dec(Col).AnchoCampo = WidthBoolean
                 Case TadDate, TadDate1
                     .Columns(Col).NumberFormat = FormatoFechas
                      If Vect_Dec(Col).AnchoCampo < WidthDate Then Vect_Dec(Col).AnchoCampo = WidthDate
                      Vect_Dec(Col).SumaTotal = FechaSistema
                 Case TadByte
                     .Columns(Col).NumberFormat = "##0"
                     .Columns(Col).Alignment = dbgRight
                      If Vect_Dec(Col).AnchoCampo < WidthByte Then Vect_Dec(Col).AnchoCampo = WidthByte
                 Case TadInteger
                     .Columns(Col).NumberFormat = "##0"
                     .Columns(Col).Alignment = dbgRight
                      If Vect_Dec(Col).AnchoCampo < WidthInteger Then Vect_Dec(Col).AnchoCampo = WidthInteger
                 Case TadLong
                     .Columns(Col).NumberFormat = "##0"
                     .Columns(Col).Alignment = dbgRight
                      If Vect_Dec(Col).AnchoCampo < WidthLong Then Vect_Dec(Col).AnchoCampo = WidthLong
                 Case TadSingle
                     .Columns(Col).NumberFormat = "##0.00%"
                     .Columns(Col).Alignment = dbgRight
                      If Vect_Dec(Col).AnchoCampo < WidthSingle Then Vect_Dec(Col).AnchoCampo = WidthSingle
                 Case TadDouble, TadCurrency
                     .Columns(Col).Alignment = dbgRight
                     .Columns(Col).NumberFormat = "#,##0." & Trim(String$(Vect_Dec(Col).CantDec, "0"))
                      If Dtas.Recordset.Fields(Col).Type = TadDouble Then
                        If Vect_Dec(Col).AnchoCampo < WidthDouble Then Vect_Dec(Col).AnchoCampo = WidthDouble
                      Else
                        If Vect_Dec(Col).AnchoCampo < WidthCurrency Then Vect_Dec(Col).AnchoCampo = WidthCurrency
                      End If
                 Case TadText
                     'Determinamos el ancho maximo mas abajo dependendo de la cantidad de caracteres por celda
                      AnchoMax = Dtas.Recordset.Fields(Col).DefinedSize
                      If AnchoMax > 50 Then AnchoMax = 50
                      If EsCampoCorto Then
                         WidthText = DBGMalla.Parent.TextWidth(String$(AnchoMax, ".") & " ")
                      Else
                         WidthText = DBGMalla.Parent.TextWidth(String$(AnchoMax, "H") & " ")
                      End If
                      If Vect_Dec(Col).AnchoCampo < WidthText Then Vect_Dec(Col).AnchoCampo = WidthText
                 Case Else
                      If Dtas.Recordset.Fields(Col).DefinedSize <= 50 Then
                         AnchoMax = Dtas.Recordset.Fields(Col).DefinedSize
                      Else
                         AnchoMax = 50
                      End If
                      WidthText = DBGMalla.Parent.TextWidth(String$(AnchoMax, "H") & " ")
                      If Vect_Dec(Col).AnchoCampo < WidthText Then Vect_Dec(Col).AnchoCampo = WidthText
               End Select
          End With
      Next Col

    'Recorremos cada columna y le asignamos el ancho
     For Col = 0 To CantCampos - 1
         DBGMalla.Columns(Col).width = Vect_Dec(Col).AnchoCampo
         'MsgBox DBGMalla.Columns(Col).width
     Next
     
     DBGMalla.Refresh
    'restablecemos la fuente del DataGrid y el scaleMode
'     Set DBGMalla.Parent.Font = saveFont
'     DBGMalla.Parent.ScaleMode = oldScaleMode
     Dtas.Caption = "Registros: " & Format$(Dtas.Recordset.RecordCount, "#,##0") & "."
   'If PresentarEsperar Then
     'Progreso_Final
  End If  'fin
  
  If Len(NombreFile) > 1 Then
     DatosFile = SQuerys
     DatosFile = Replace(DatosFile, "FROM", vbCrLf & "FROM")
     DatosFile = Replace(DatosFile, "WHERE", vbCrLf & "WHERE")
     DatosFile = Replace(DatosFile, "AND", vbCrLf & "AND")
     DatosFile = Replace(DatosFile, "OR ", vbCrLf & "OR ")
     DatosFile = Replace(DatosFile, "SET", vbCrLf & "SET")
     DatosFile = Replace(DatosFile, "GROUP BY", vbCrLf & "GROUP BY")
     DatosFile = Replace(DatosFile, "ORDER BY", vbCrLf & "ORDER BY")
     DatosFile = Replace(DatosFile, "HAVING", vbCrLf & "HAVING")
     DatosFile = Replace(DatosFile, "VALUES", vbCrLf & "VALUES" & vbCrLf)
     NumFile = FreeFile
     Open RutaSysBases & "\TEMP\" & NombreFile & ".sql" For Output As #NumFile
     Print #NumFile, DatosFile
     Close #NumFile
  End If
  
  If Presentar Then DBGMalla.Visible = True
  RatonNormal
'''  If Is_Prog_Bar Then Prog_Bar.Value = Prog_Bar.Max
Exit Sub
Errorhandler:
    RatonNormal
    DBGMalla.Visible = True
    MsgBox ("Error #" & CStr(Err.Number) & " " & Err.Description)
    Err.Clear    ' Limpia el error.
    Dtas.Caption = "Registros: 0"
    Exit Sub
End Sub

Public Function AnchoTipoCampoTexto(AdoTipo As ADODB.Field) As Single
Dim ch As String
  With AdoTipo
    Select Case .Type
      Case TadDate, TadDate1: ch = CadDate
      Case TadTime:     ch = CadTime
      Case TadBoolean:  ch = CadBoolean
      Case TadByte:     ch = CadByte
      Case TadInteger:  ch = CadInteger
      Case TadLong:     ch = CadLong
      Case TadSingle:   ch = CadSingle
      Case TadDouble:   ch = CadDouble
      Case TadCurrency: ch = CadCurrency
      Case TadText:     ch = String(.DefinedSize, "F") & " "
      Case Else:        ch = String(10, "F") & " "
    End Select
  End With
If ch = "" Then ch = Ninguno
AnchoTipoCampoTexto = Len(ch)
End Function

Public Function CompilarRUC_CI(CadSQL As String) As String
Dim StrSQL As String
Dim Indc As Integer
 If ContadorRUCCI <= 0 Then ContadorRUCCI = 1
'MsgBox CadSQL
 If (CadSQL = Ninguno) Or (Val(CadSQL) = 0) Then
    StrSQL = Format(ContadorRUCCI, "00000000")
    ContadorRUCCI = ContadorRUCCI + 1
 ElseIf Len(CadSQL) < 8 Then
    StrSQL = CadSQL & String(8 - Len(CadSQL), "0")
 Else
    StrSQL = ""
    If Len(CadSQL) > 0 Then
       StrSQL = ""
       For Indc = 1 To Len(CadSQL)
           If Mid(CadSQL, Indc, 1) <> "-" Then StrSQL = StrSQL & Mid(CadSQL, Indc, 1)
       Next Indc
    End If
 End If
 If (Len(StrSQL) > 10) And (Mid(StrSQL, Len(StrSQL) - 2, 3) = "000") Then StrSQL = Mid(StrSQL, 1, Len(StrSQL) - 3)
 If Val(StrSQL) <= 0 Then
    StrSQL = Format(ContadorRUCCI, "00000000")
    ContadorRUCCI = ContadorRUCCI + 1
 End If
 CompilarRUC_CI = StrSQL
End Function

Public Function CompilarSQL(CadSQL As String)
Dim StrSQL As String
Dim Indc As Integer
 StrSQL = CadSQL
 If SQL_Server Then
    If Len(CadSQL) > 0 Then
       StrSQL = ""
       For Indc = 1 To Len(CadSQL)
           If Mid(CadSQL, Indc, 1) <> "#" Then
              If Mid(CadSQL, Indc, 1) = "&" Then
                 StrSQL = StrSQL & "+"
              Else
                 StrSQL = StrSQL & Mid(CadSQL, Indc, 1)
              End If
           ElseIf Mid(CadSQL, Indc, 1) = "#" Then
                 StrSQL = StrSQL & "'"
           End If
       Next Indc
    Else
       StrSQL = ""
    End If
    CadSQL = StrSQL
    If UCase(Mid(CadSQL, 1, 6)) = "DELETE" And Len(CadSQL) > 0 Then
       StrSQL = ""
       For Indc = 1 To Len(CadSQL)
           If Mid(CadSQL, Indc, 1) <> "*" Then
              StrSQL = StrSQL & Mid(CadSQL, Indc, 1)
           End If
       Next Indc
    End If
    StrSQL = Replace(StrSQL, "Mid(", "SUBSTRING(")
    StrSQL = Replace(StrSQL, "MID(", "SUBSTRING(")
    StrSQL = Replace(StrSQL, "mid(", "SUBSTRING(")
    StrSQL = Replace(StrSQL, "UCase(", "UPPER(")
    StrSQL = Replace(StrSQL, "Ucase(", "UPPER(")
    StrSQL = Replace(StrSQL, "UCASE(", "UPPER(")
    StrSQL = Replace(StrSQL, "CSTR(", "STR(")
    StrSQL = Replace(StrSQL, "CStr(", "STR(")
    StrSQL = Replace(StrSQL, "cstr(", "STR(")
    StrSQL = Replace(StrSQL, "False", "0")
    StrSQL = Replace(StrSQL, "True", "1")
    StrSQL = Replace(StrSQL, "FALSE", "0")
    StrSQL = Replace(StrSQL, "TRUE", "1")
    StrSQL = Replace(StrSQL, "false", "0")
    StrSQL = Replace(StrSQL, "true", "1")
 End If
 CompilarSQL = StrSQL
End Function

Public Sub ConectarAdodc(AdoBase As Adodc)
  AdoBase.ConnectionString = AdoStrCnn
End Sub

Public Sub ConectarAdoExecute(SQLQuery As String, _
                              Optional RegSN As Boolean)
Dim AdoCon1 As ADODB.Connection
Dim IdTime As Long
  RatonReloj
 'Consultamos las cuentas de la tabla
  SQLQuery = CompilarSQL(SQLQuery)
 'MsgBox SQLQuery
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.Open AdoStrCnn
  AdoCon1.Execute SQLQuery, RegAfectados, adCmdText
  AdoCon1.Close
  RatonNormal
  If RegSN Then MsgBox "Registros Afectados: " & Format(RegAfectados, "#,##0")
End Sub

Public Sub CentrarForm(Forms As Form)
Dim PosSup, PosIzq As Single
  'Centrar el formulario
  'MsgBox MDI_X_Max & vbCrLf & MDI_Y_Max
   If MDI_X_Max > 0 And MDI_Y_Max > 0 Then
      PosIzq = ((MDI_X_Max - Forms.width) / 2)
      PosSup = ((MDI_Y_Max - Forms.Height) / 2)
   Else
      PosIzq = ((Screen.width - Forms.width) / 2)
      PosSup = ((Screen.Height - Forms.Height) / 2)
   End If
   If Forms.BorderStyle = 0 Then PosSup = PosSup - 200
   
   If PosIzq < 0 Then PosIzq = 0
   If PosSup < 0 Then PosSup = 0
   Forms.Left = PosIzq: Forms.Top = PosSup - 400
End Sub

Public Sub Control_Procesos(TipoTrans As String, Tarea As String, Optional Proceso As String)
'No hace nada
End Sub

' Subrutina que copia el archivo
Public Sub Copiar_Archivo(ByVal Origen As String, ByVal Destino As String)
Dim t_Op As SHFILEOPSTRUCT
    With t_Op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = Origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With

    ' Se ejecuta la función Api pasandole la estructura
    SHFileOperation t_Op
End Sub

Public Sub Conectar_Base_Datos()
Dim CarBase As String
Dim Conexion_Temp As String
Dim Nombre_Base_Datos_SQL As String
Dim CarIni As Integer
Dim CarFin As Integer

  'Leemos la cadena de conexion
   RutaGeneraFile = RutaSistema & "\ConectarDB.ini"
   RutaEmpresa = RutaSistema & "\EMPRESA\"
   Conexion_Temp = ""
   Nombre_Base_Datos_SQL = ""
   Nombre_Base_SQL = ""
   
   NumFile = FreeFile
   Open RutaGeneraFile For Input As #NumFile
   Do While Not EOF(NumFile)
      CarBase = Input(1, #NumFile) ' Obtiene un carácter.
      Conexion_Temp = Conexion_Temp & CarBase
   Loop
   Close #NumFile
   
   If InStr(1, Conexion_Temp, "SQLServer = SI") Then SQL_Server = True Else SQL_Server = False
   If SQL_Server Then
      CarIni = InStr(1, Conexion_Temp, "<SQLServer = SI>") + 18
      CarFin = InStr(1, Conexion_Temp, "</SQLServer>") - 2
      Conexion_Temp = Mid$(Conexion_Temp, CarIni, CarFin - CarIni)
      
      CarIni = InStr(1, Conexion_Temp, "Initial Catalog=")
      Nombre_Base_Datos_SQL = Mid$(Conexion_Temp, CarIni, Len(Conexion_Temp))
      CarFin = InStr(1, Nombre_Base_Datos_SQL, ";")
      Nombre_Base_SQL = Mid$(Nombre_Base_Datos_SQL, 17, CarFin - 17)
      PathEmpresa = ""
   Else
      CarIni = InStr(1, Conexion_Temp, "<Access = SI>") + 15
      CarFin = InStr(1, Conexion_Temp, "</Access>") - 2
      Conexion_Temp = Mid$(Conexion_Temp, CarIni, CarFin - CarIni)
      Conexion_Temp = Replace(Conexion_Temp, "Path\", RutaSistema & "\EMPRESA\")
      
      CarIni = InStr(1, Conexion_Temp, "Data Source=")
      Nombre_Base_Datos_SQL = Mid$(Conexion_Temp, CarIni, Len(Conexion_Temp))
      CarFin = InStr(1, Nombre_Base_Datos_SQL, ";")
      Nombre_Base_SQL = Mid$(Nombre_Base_Datos_SQL, 13, CarFin - 13)
      PathEmpresa = Nombre_Base_SQL
   End If
   AdoStrCnn = Conexion_Temp
  'MsgBox PathEmpresa & vbCrLf & AdoStrCnn
End Sub

Public Sub SelectDBCombo(DBCombos As DataCombo, _
                         DataSQL As Adodc, _
                         SQLs As String, _
                         NombreCampo As String, _
                         Optional Final As Boolean)
  If SQLs <> "" Then
     SQLs = CompilarSQL(SQLs)
    'MsgBox SQLs
     DataSQL.RecordSource = SQLs
     DataSQL.Refresh
     DBCombos.ListField = DataSQL.Recordset.Fields(NombreCampo).Name
     If DataSQL.Recordset.RecordCount > 0 Then
         If Final Then DataSQL.Recordset.MoveLast
         DBCombos.Text = DataSQL.Recordset.Fields(NombreCampo)
         DBCombos.SelStart = 0
         DBCombos.SelLength = Len(DBCombos.Text)
     Else
         DBCombos.Text = "No existen datos."
     End If
  End If
End Sub

Public Sub SelectDBList(DBLists As DataList, _
                        DataSQL As Adodc, _
                        SQLs As String, _
                        NombreCampo As String)
If SQLs <> "" Then
   SQLs = CompilarSQL(SQLs)
   'MsgBox SQLs
   DataSQL.RecordSource = SQLs
   DataSQL.Refresh
   DBLists.ListField = DataSQL.Recordset.Fields(NombreCampo).Name
   If DataSQL.Recordset.RecordCount > 0 Then
      DBLists.Text = DataSQL.Recordset.Fields(NombreCampo)
   End If
End If
End Sub

Public Function BuscarFecha(FechaStr As String) As String
'dd/mm/yyyy
  If IsNumeric(FechaStr) Then
     If SQL_Server Then
        BuscarFecha = Format$(FechaSistema, "YYYYMMDD")
     Else
        BuscarFecha = Format$(FechaSistema, "mm/dd/yyyy")
     End If
     'MsgBox "Fecha Incorrecta"
  Else
     If SQL_Server Then
        BuscarFecha = Format$(FechaStr, "YYYYMMDD")
     Else
        BuscarFecha = Format$(FechaStr, "mm/dd/yyyy")
     End If
  End If
End Function

Public Sub PresionoEnter(KeyCode)
  If KeyCode = vbKeyReturn Then Pulsar_Tecla (vbKeyTab)    ' SendKeys "{TAB}", False
End Sub

Sub Pulsar_Tecla(Tecla As Long)
    Call keybd_event(Tecla, 0, 0, 0)
    Call keybd_event(Tecla, 0, KEYEVENTF_KEYUP, 0)
End Sub

Public Sub SiguienteControl()
   Pulsar_Tecla (vbKeyTab)
End Sub

Public Sub Keys_Especiales(KeyShift As Integer)
  ShiftDown = (KeyShift And vbShiftMask) > 0
  AltDown = (KeyShift And vbAltMask) > 0
  CtrlDown = (KeyShift And vbCtrlMask) > 0
End Sub

