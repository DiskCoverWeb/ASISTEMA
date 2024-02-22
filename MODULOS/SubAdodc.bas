Attribute VB_Name = "SubAdodc"
Option Explicit
'SetAdoAddNew "Nombre_Tabla"
'SetAdoFields "", 0
'SetAdoUpdate
'½
'''  If SQL_Server Then
'''     sSQL = "UPDATE Transacciones " _
'''          & "SET Cheq_Dep = TB.Cheq_Dep " _
'''          & "FROM Transacciones As T,Trans_Bancos As TB "
'''  Else
'''     sSQL = "UPDATE Transacceiones As T,Trans_Bancos As TB " _
'''          & "SET T.Cheq_Dep = TB.Cheq_Dep "
'''  End If
'''  sSQL = sSQL & "WHERE T.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''       & "AND T.TP = TB.TP " _
'''       & "AND T.Numero = TB.Numero " _
'''       & "AND T.Fecha = TB.Fecha " _
'''       & "AND T.Item = TB.Item " _
'''       & "AND T.Cta = TB.Cta "
'''  Ejecutar_SQL_SP sSQL

'''Dim AdoAuxDB As ADODB.Recordset
'''  sSQL = "SELECT * " _
'''       & "FROM " & NombreTabla & " " _
'''       & "WHERE 1 = 0 "
'''  Select_AdoDB AdoAuxDB, sSQL
'''  AdoAuxDB.Close

Public Function Crear_FN_SP(RutaFile As String) As String
Dim Si_Crear As Boolean
Dim NumFile As Long
Dim Si_File_I As Long
Dim Si_File_F As Long
Dim TextFile As String
Dim LineFile As String
Dim Results As String
  
  RatonReloj
  Si_File_I = 0
  Si_File_F = 0
  Nombre_FN_SP = ""
  TextFile = ""
  Si_File_I = InStr(RutaFile, "dbo.")
  Si_Crear = False
  If Len(RutaFile) > 1 Then
     Results = Dir$(RutaFile)
     If Results <> "" Then
        NumFile = FreeFile
        Open RutaFile For Input As #NumFile
        Do While Not EOF(NumFile)
           Line Input #NumFile, LineFile
           If InStr(LineFile, "CREATE FUNCTION") > 0 And Not Si_Crear Then
              Si_Crear = True
              Si_File_F = InStr(RutaFile, ".UserDefinedFunction.")
           End If
           If InStr(LineFile, "CREATE PROCEDURE") > 0 And Not Si_Crear Then
              Si_Crear = True
              Si_File_F = InStr(RutaFile, ".StoredProcedure.")
           End If
           If Si_Crear Then TextFile = TextFile & LineFile & vbCrLf
        Loop
        Close #NumFile
     End If
  End If
  RatonNormal
  Nombre_FN_SP = MidStrg(RutaFile, Si_File_I, Si_File_F - Si_File_I)
  TextFile = TrimStrg(MidStrg(TextFile, 1, Len(TextFile) - 4))
  Crear_FN_SP = TrimStrg(TextFile)
End Function

Public Sub Generar_File_SQL(NombreFile As String, SQLQuery As String)
Dim DatosFile As String
Dim NumFile As Long

  If Len(NombreFile) > 1 Then
     RatonReloj
     DatosFile = SQLQuery
     DatosFile = Replace(DatosFile, "SELECT ", vbCrLf & "SELECT ")
     DatosFile = Replace(DatosFile, "UNION ", vbCrLf & "UNION ")
     DatosFile = Replace(DatosFile, "FROM ", vbCrLf & "FROM ")
     DatosFile = Replace(DatosFile, "WHERE ", vbCrLf & "WHERE ")
     DatosFile = Replace(DatosFile, "AND ", vbCrLf & "AND ")
     DatosFile = Replace(DatosFile, "OR ", vbCrLf & "OR ")
     DatosFile = Replace(DatosFile, "SET ", vbCrLf & "SET ")
     DatosFile = Replace(DatosFile, "GROUP BY ", vbCrLf & "GROUP BY ")
     DatosFile = Replace(DatosFile, "ORDER BY ", vbCrLf & "ORDER BY ")
     DatosFile = Replace(DatosFile, "HAVING ", vbCrLf & "HAVING ")
     DatosFile = Replace(DatosFile, "VALUES ", vbCrLf & "VALUES" & vbCrLf)
'     DatosFile = Replace(DatosFile, "NULL,", "NULL," & vbCrLf)
     NumFile = FreeFile
     Open RutaSysBases & "\TEMP\" & NombreFile & ".sql" For Output As #NumFile
     Print #NumFile, DatosFile
     Close #NumFile
     RatonNormal
  End If
End Sub

Public Sub Datos_IESS(FechaMes As String)
Dim RegIESS As ADODB.Recordset
Dim sSQL1 As String
 If IsDate(FechaMes) Then
    sSQL1 = "SELECT Codigo, Porc " _
          & "FROM Tabla_Por_ICE_IVA " _
          & "WHERE Codigo IN ('IESS_Per','IESS_Pat','IESS_ExtC','Sueldo_Bas','Canasta_Ba') " _
          & "AND Fecha_Inicio <= #" & BuscarFecha(FechaMes) & "# " _
          & "AND Fecha_Final >= #" & BuscarFecha(FechaMes) & "# " _
          & "ORDER BY Codigo "
    Select_AdoDB RegIESS, sSQL1
    With RegIESS
     If .RecordCount > 0 Then
         Do While Not .EOF
            Select Case .fields("Codigo")
              Case "IESS_Per": IESS_Per = .fields("Porc") / 100
              Case "IESS_Pat": IESS_Pat = .fields("Porc") / 100
              Case "IESS_ExtC": IESS_Ext = .fields("Porc") / 100
              Case "Sueldo_Bas": Sueldo_Basico = .fields("Porc")
              Case "Canasta_Ba": Canasta_Basica = .fields("Porc")
            End Select
           .MoveNext
         Loop
     End If
    End With
    RegIESS.Close
 Else
    IESS_Per = 0
    IESS_Pat = 0
    IESS_Ext = 0
    Sueldo_Basico = 0
    Canasta_Basica = 0
 End If
End Sub

Public Sub Email_Memo(para As Adodc, _
                      Datas As Adodc, _
                      NumeroMemo As Long)
Dim SizeLetra As Integer
Dim TextoMemo As String
Dim TextoTemp As String

On Error GoTo Errorhandler
SizeLetra = 10
Mensajes = "Seguro de Envial Memo No. " & NumeroMemo & vbCrLf
Titulo = "ENVIOS POR MAIL"
If BoxMensaje = vbYes Then
   RatonReloj
 ' Iniciamos la impresion
   sSQL = "SELECT A.Nombre_Completo,C.Cliente,C.Email,TM.* " _
        & "FROM Accesos As A, Clientes As C, Trans_Memos As TM " _
        & "WHERE TM.Numero = " & NumeroMemo & " " _
        & "AND TM.Item = '" & NumEmpresa & "' " _
        & "AND A.Codigo = TM.CodigoU " _
        & "AND C.Codigo = TM.Codigo "
   Select_Adodc Datas, sSQL
   DataAnchoCampos InicioX, Datas, SizeLetra, TipoCourier, Orientacion_Pagina
    With Datas.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         If Len(.fields("Nombre_Completo")) > 1 Then
            Codigo1 = .fields("Nombre_Completo")
            If para.Recordset.RecordCount > 0 Then
               para.Recordset.MoveFirst
               para.Recordset.Find ("Cliente = '" & Codigo1 & "' ")
               If Not para.Recordset.EOF Then Codigo2 = para.Recordset.fields("Email")
            End If
         End If
         
         TMail.Asunto = .fields("Asunto")
         TMail.Mensaje = String$(20, " ") & "M E M O R A N D O" & vbCrLf _
                       & String$(20, " ") & "-----------------" & vbCrLf _
                       & "Fecha: " & .fields("Fecha") _
                       & String$(50, " ") & "Numero No. " _
                       & Format$(Day(.fields("Fecha")), "00") & Format$(Month(.fields("Fecha")), "00") & "-" & Format$(.fields("Numero"), "000000") & vbCrLf _
                       & "DE: " & UCaseStrg(.fields("Nombre_Completo")) & ", " _
                       & "Email: " & Codigo2 & vbCrLf _
                       & "PARA: " & UCaseStrg(.fields("Cliente")) & ", " _
                       & "Email: " & .fields("Email") & vbCrLf _
                       & "ATENCION: " & .fields("Atencion") & vbCrLf _
                       & String$(70, "_") & vbCrLf _
                       & .fields("Texto_Memo") & vbCrLf & vbCrLf _
                       & "Atentamente," & vbCrLf & vbCrLf _
                       & .fields("Nombre_Completo") & vbCrLf _
                       & Empresa
          TMail.para = .fields("Email")
          If EsUnEmail(TMail.para) Then FEnviarCorreos.Show 1
          If Len(.fields("CC1")) > 1 Then
             Codigo1 = .fields("CC1")
             If para.Recordset.RecordCount > 0 Then
                para.Recordset.MoveFirst
                para.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
                If Not para.Recordset.EOF Then
                   TMail.para = para.Recordset.fields("Email")
                   If EsUnEmail(TMail.para) Then FEnviarCorreos.Show 1
                End If
             End If
          End If
          If Len(.fields("CC2")) > 1 Then
             Codigo1 = .fields("CC2")
             If para.Recordset.RecordCount > 0 Then
                para.Recordset.MoveFirst
                para.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
                If Not para.Recordset.EOF Then
                   TMail.para = para.Recordset.fields("Email")
                   If EsUnEmail(TMail.para) Then FEnviarCorreos.Show 1
                End If
             End If
          End If
          MsgBox "Envio Exitoso por correo electrónico"
     Else
         MsgBox "No hay Datos para Enviar"
     End If
    End With
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

Public Sub Imprimir_Memo(para As Adodc, _
                         Datas As Adodc, _
                         NumeroMemo As Long)
Dim SizeLetra As Integer
Dim TextoMemo As String
Dim TextoTemp As String
On Error GoTo Errorhandler
SizeLetra = 10
Mensajes = "Seguro de Imprimir el Memo No. " & NumeroMemo & vbCrLf
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
   RatonReloj
   InicioX = 2: InicioY = 0.1
   Pagina = 1
   LimiteAlto = LimiteAlto - 1
 ' Iniciamos la impresion
   sSQL = "SELECT A.Nombre_Completo,C.Cliente,TM.* " _
        & "FROM Accesos As A, Clientes As C, Trans_Memos As TM " _
        & "WHERE TM.Numero = " & NumeroMemo & " " _
        & "AND TM.Item = '" & NumEmpresa & "' " _
        & "AND A.Codigo = TM.CodigoU " _
        & "AND C.Codigo = TM.Codigo "
   Select_Adodc Datas, sSQL
   DataAnchoCampos InicioX, Datas, SizeLetra, TipoCourier, Orientacion_Pagina
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     Codigo1 = .fields("CC1")
     Codigo2 = .fields("CC2")
     Cadena1 = Ninguno: Cadena2 = Ninguno
     If para.Recordset.RecordCount > 0 Then
        para.Recordset.MoveFirst
        para.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
        If Not para.Recordset.EOF Then Cadena1 = para.Recordset.fields("Cliente")
     End If
     If para.Recordset.RecordCount > 0 Then
        para.Recordset.MoveFirst
        para.Recordset.Find ("Codigo = '" & Codigo2 & "' ")
        If Not para.Recordset.EOF Then Cadena2 = para.Recordset.fields("Cliente")
     End If
     TextoMemo = .fields("Texto_Memo")
     Encabezado InicioX, 18.5
     Printer.FontSize = 22
     Printer.FontBold = True
     PrinterTexto 6.2, PosLinea, "M E M O R A N D O"
     PosLinea = PosLinea + 1.3
     Printer.FontSize = SizeLetra
     PrinterTexto 14, PosLinea, "NUMERO:"
     PrinterTexto InicioX, PosLinea, "FECHA:"
     Printer.FontBold = False
     PrinterTexto 15.5, PosLinea, Format$(Day(.fields("Fecha")), "00") & Format$(Month(.fields("Fecha")), "00") & "-" & Format$(.fields("Numero"), "000000")
     PrinterTexto InicioX + 1.5, PosLinea, .fields("Fecha")
     PosLinea = PosLinea + 0.5
     Imprimir_Linea_H PosLinea, InicioX, 18.5, Negro, True
     PosLinea = PosLinea + 0.1
     
     Printer.FontBold = True
     PrinterTexto InicioX, PosLinea, "DE:"
     Printer.FontBold = False
     PrinterTexto InicioX + 2, PosLinea, UCaseStrg(.fields("Nombre_Completo"))
     PosLinea = PosLinea + 0.5

     Printer.FontBold = True
     PrinterTexto InicioX, PosLinea, "PARA:"
     Printer.FontBold = False
     PrinterTexto InicioX + 2, PosLinea, .fields("Cliente")
     PosLinea = PosLinea + 0.5
     
     If Cadena1 <> Ninguno Then
        Printer.FontBold = True
        PrinterTexto InicioX, PosLinea, "C.C.:"
        Printer.FontBold = False
        PrinterTexto InicioX + 2, PosLinea, Cadena1
        PosLinea = PosLinea + 0.5
     End If
     
     If Cadena2 <> Ninguno Then
        Printer.FontBold = True
        PrinterTexto InicioX, PosLinea, "C.C.:"
        Printer.FontBold = False
        PrinterTexto InicioX + 2, PosLinea, Cadena2
        PosLinea = PosLinea + 0.5
     End If
     
     Printer.FontBold = True
     PrinterTexto InicioX, PosLinea, "ASUNTO:"
     Printer.FontBold = False
     PrinterTexto InicioX + 2, PosLinea, .fields("Asunto")
     PosLinea = PosLinea + 0.5
     
     Printer.FontBold = True
     PrinterTexto InicioX, PosLinea, "ATENCION:"
     Printer.FontBold = False
     PrinterTexto InicioX + 2, PosLinea, .fields("Atencion")
     PosLinea = PosLinea + 0.5
     Printer.FontSize = 12
     Printer.FontBold = True
     PrinterTexto InicioX, PosLinea, "D E T A L L E:"
     PosLinea = PosLinea + 0.5
     Printer.FontBold = False
     Imprimir_Linea_H PosLinea, InicioX, 18.5, Negro, True
     PosLinea = PosLinea + 0.1
     Printer.FontSize = SizeLetra
     TextoTemp = ""
     For I = 1 To Len(.fields("Texto_Memo"))
         TextoMemo = MidStrg(.fields("Texto_Memo"), I, 1)
         Select Case Asc(TextoMemo)
           Case 13
                  TextoTemp = TextoTemp & " "
           Case 10
                  TextoTemp = TextoTemp & " " & vbCrLf
                  PrinterTexto InicioX, PosLinea, TrimStrg(TextoTemp), , 16
                  TextoTemp = ""
                  PosLinea = PosLinea + 0.4
                  If PosLinea >= LimiteAlto Then
                     Printer.NewPage
                     Encabezado InicioX, 18.5
                     Printer.FontSize = SizeLetra
                     PrinterTexto 14, PosLinea, "NUMERO:"
                     PrinterTexto InicioX, PosLinea, "FECHA:"
                     Printer.FontBold = False
                     PrinterTexto 15.5, PosLinea, Format$(Day(.fields("Fecha")), "00") & Format$(Month(.fields("Fecha")), "00") & "-" & Format$(.fields("Numero"), "000000")
                     PrinterTexto InicioX + 1.5, PosLinea, .fields("Fecha")
                     PosLinea = PosLinea + 0.5
                     Imprimir_Linea_H PosLinea, InicioX, 18.5, Negro, True
                     PosLinea = PosLinea + 0.1
                     Printer.FontSize = SizeLetra
                  End If
           Case Else
                TextoTemp = TextoTemp & TextoMemo
                If Printer.TextWidth(TextoTemp) >= 16 Then
                   If MidStrg(.fields("Texto_Memo"), I + 1, 1) <> " " Then
                      Do While TextoMemo <> " "
                         TextoTemp = MidStrg(TextoTemp, 1, Len(TextoTemp) - 1)
                         I = I - 1
                         TextoMemo = MidStrg(.fields("Texto_Memo"), I, 1)
                      Loop
                   End If
                   PrinterTexto InicioX, PosLinea, TrimStrg(TextoTemp), , 16
                   TextoTemp = ""
                   PosLinea = PosLinea + 0.4
                   If PosLinea >= LimiteAlto Then
                      Printer.NewPage
                      Encabezado InicioX, 18.5
                      Printer.FontSize = SizeLetra
                      PrinterTexto 14, PosLinea, "NUMERO:"
                      PrinterTexto InicioX, PosLinea, "FECHA:"
                      Printer.FontBold = False
                      PrinterTexto 15.5, PosLinea, Format$(Day(.fields("Fecha")), "00") & Format$(Month(.fields("Fecha")), "00") & "-" & Format$(.fields("Numero"), "000000")
                      PrinterTexto InicioX + 1.5, PosLinea, .fields("Fecha")
                      PosLinea = PosLinea + 0.5
                      Imprimir_Linea_H PosLinea, InicioX, 18.5, Negro, True
                      PosLinea = PosLinea + 0.1
                      Printer.FontSize = SizeLetra
                   End If
                End If
         End Select
     Next I
     PosLinea = PosLinea + 0.5
     PrinterTexto 2, PosLinea, "Atentamente,"
     PosLinea = PosLinea + 2
     Printer.FontBold = True
     PrinterTexto 2, PosLinea, .fields("Nombre_Completo")
     PosLinea = PosLinea + 0.5
     PrinterTexto 2, PosLinea, Empresa
 Else
     MsgBox "No hay Datos para Imprimir"
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

Public Function CompilarSQL(CadSQL As String) As String
Dim StrSQL As String
Dim Indc As Long
Dim Fecha_SQL As String
Dim Inic_Fecha As Boolean
 Fecha_SQL = ""
 Inic_Fecha = False
 StrSQL = CadSQL
 If SQL_Server Then
    If Len(CadSQL) > 0 Then
       StrSQL = ""
       For Indc = 1 To Len(CadSQL)
           If MidStrg(CadSQL, Indc, 1) <> "#" Then
              If MidStrg(CadSQL, Indc, 1) = "&" Then
                 StrSQL = StrSQL & "+"
              Else
                 StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
              End If
           ElseIf MidStrg(CadSQL, Indc, 1) = "#" Then
                 StrSQL = StrSQL & "'"
                 Inic_Fecha = Not Inic_Fecha
           End If
           If Inic_Fecha Then Fecha_SQL = Fecha_SQL & MidStrg(CadSQL, Indc, 1)
       Next Indc
    Else
       StrSQL = ""
    End If
    CadSQL = StrSQL
    If UCaseStrg(MidStrg(CadSQL, 1, 6)) = "DELETE" And Len(CadSQL) > 0 Then
       StrSQL = ""
       For Indc = 1 To Len(CadSQL)
           If MidStrg(CadSQL, Indc, 1) <> "*" Then
              StrSQL = StrSQL & MidStrg(CadSQL, Indc, 1)
           End If
       Next Indc
    End If
    StrSQL = Replace(StrSQL, "MidStrg(", "SUBSTRING(")
    StrSQL = Replace(StrSQL, "UCaseStrg(", "UPPER(")
    StrSQL = Replace(StrSQL, "LeftStrg(", "LTRIM(")
    StrSQL = Replace(StrSQL, "RightStrg(", "RTRIM(")
 Else
    StrSQL = Replace(StrSQL, "MidStrg(", "MidStrg(")
    StrSQL = Replace(StrSQL, "UCaseStrg(", "UCase$(")
    StrSQL = Replace(StrSQL, "LeftStrg(", "Ltrim$(")
    StrSQL = Replace(StrSQL, "RightStrg(", "RTrim$(")
 End If
 StrSQL = Replace(StrSQL, "CSTR(", "STR(")
 StrSQL = Replace(StrSQL, "CStr(", "STR(")
 StrSQL = Replace(StrSQL, "False", "0")
 StrSQL = Replace(StrSQL, "True", "1")
 StrSQL = Replace(StrSQL, "false", "0")
 StrSQL = Replace(StrSQL, "true", "1")
 If SQL_Server Then StrSQL = Replace(StrSQL, "#", "'")
 CompilarSQL = StrSQL
End Function

Public Function FieldType(Campo As String, _
                          IntType As Integer) As String
Select Case IntType
  Case adBoolean: FieldType = "adBoolean"
  Case adTinyInt: FieldType = "adTinyInt"
  Case adSmallInt: FieldType = "adSmallInt"
  Case adInteger: FieldType = "adInteger"
  Case adDouble: FieldType = "adDouble"
  Case adBigInt: FieldType = "adBigInt"
  Case adUnsignedTinyInt: FieldType = "adUnsignedTinyInt"
  Case adUnsignedSmallInt: FieldType = "adUnsignedSmallInt"
  Case adUnsignedInt: FieldType = "adUnsignedInt"
  Case adUnsignedBigInt: FieldType = "adUnsignedBigInt"
  Case adSingle: FieldType = "adSingle"
  Case adCurrency: FieldType = "adCurrency"
  Case adDecimal: FieldType = "adDecimal"
  Case adNumeric: FieldType = "adNumeric"
  Case adUserDefined: FieldType = "adUserDefined"
  Case adVariant: FieldType = "adVariant"
  Case adGUID: FieldType = "adGUID"
  Case adDate: FieldType = "adDate"
  Case adDBDate: FieldType = "adDBDate"
  Case adDBTime: FieldType = "adDBTime"
  Case adDBTimeStamp: FieldType = "adDBTimeStamp"
  Case adBSTR: FieldType = "adBSTR"
  Case adChar: FieldType = "adChar"
  Case adVarChar: FieldType = "adVarChar"
  Case adLongVarChar: FieldType = "adLongVarChar"
  Case adWChar: FieldType = "adWChar"
  Case adVarWChar: FieldType = "adVarWChar"
  Case adLongVarWChar: FieldType = "adLongVarWChar"
  Case adBinary: FieldType = "adBinary"
  Case adVarBinary: FieldType = "adVarBinary"
  Case adLongVarBinary: FieldType = "adLongVarBinary"
  Case Else: FieldType = "-x-"
End Select
End Function

Public Function TypeField(StrType As String) As Integer
Dim PIni As Long
Dim CampoTypeField As Integer
Dim StrTypeTemp As String

  CampoTypeField = 0
  PIni = InStr(StrType, "(")
  If PIni > 0 Then StrTypeTemp = MidStrg(StrType, 1, PIni - 1) Else StrTypeTemp = StrType
  StrTypeTemp = TrimStrg(Replace(StrTypeTemp, "NULL", ""))
  If SQL_Server Then
      Select Case StrTypeTemp
        Case "BIT": CampoTypeField = adBoolean
        Case "TINYINT": CampoTypeField = adUnsignedTinyInt
        Case "SMALLINT": CampoTypeField = adSmallInt
        Case "INT": CampoTypeField = adInteger
        Case "REAL": CampoTypeField = adSingle
        Case "FLOAT": CampoTypeField = adDouble
        Case "MONEY": CampoTypeField = adCurrency
        Case "DECIMAL": CampoTypeField = adNumeric
        Case "UNIQUEIDENTIFIER": CampoTypeField = adGUID
        Case "NTEXT": CampoTypeField = adLongVarWChar
        Case "NVARCHAR": CampoTypeField = adVarWChar
        Case "DATETIME", "SMALLDATETIME": CampoTypeField = adDate
        Case Else
             CampoTypeField = adVarWChar
      End Select
      If TrimStrg(Replace(StrType, "NULL", "")) = "NVARCHAR(MAX)" Then CampoTypeField = 203
  Else
      Select Case StrTypeTemp
        Case "BIT": CampoTypeField = adBoolean
        Case "BYTE": CampoTypeField = adUnsignedTinyInt
        Case "SHORT": CampoTypeField = adSmallInt
        Case "LONG": CampoTypeField = adInteger
        Case "SINGLE": CampoTypeField = adSingle
        Case "DOUBLE": CampoTypeField = adDouble
        Case "CURRENCY": CampoTypeField = adCurrency
        Case "DECIMAL": CampoTypeField = adDecimal
        Case "INTEGER": CampoTypeField = adNumeric
        Case "GUID": CampoTypeField = adGUID
        Case "LONGTEXT": CampoTypeField = adLongVarWChar
        Case "TEXT": CampoTypeField = adVarWChar
        Case "DATETIME": CampoTypeField = adDate
        Case Else
             CampoTypeField = adVarWChar
      End Select
  End If
  TypeField = CampoTypeField
End Function

Public Function FieldTypeSQL(IntType As Integer) As String
  Select Case IntType
    Case adSmallInt: FieldTypeSQL = "SMALLINT"       '   2
    Case adInteger: FieldTypeSQL = "INT"             '   3
    Case adSingle: FieldTypeSQL = "REAL"             '   4
    Case adDouble: FieldTypeSQL = "FLOAT"            '   5
    Case adCurrency: FieldTypeSQL = "MONEY"          '   6
    Case adDate: FieldTypeSQL = "DATETIME"           '   7
    Case adBoolean: FieldTypeSQL = "BIT"             '  11
    Case adTinyInt: FieldTypeSQL = "TINYINT"         '  16
    Case adUnsignedTinyInt: FieldTypeSQL = "TINYINT" '  17
    Case adGUID: FieldTypeSQL = "UNIQUEIDENTIFIER"   '  72
    Case adNumeric: FieldTypeSQL = "DECIMAL"         ' 131
    Case adDBTime: FieldTypeSQL = "DATETIME"         ' 134
    Case adDBTimeStamp: FieldTypeSQL = "DATETIME"    ' 135
    Case adVarWChar: FieldTypeSQL = "NVARCHAR"       ' 202
    Case adLongVarWChar: FieldTypeSQL = "NVARCHAR"   ' 203
    Case Else
'        MsgBox IntType & vbCrLf & FieldType("", IntType)
         FieldTypeSQL = "-x-"
  End Select
End Function

Public Function FieldTypeAccess(IntType As Integer) As String
  Select Case IntType
    Case adBoolean: FieldTypeAccess = "BIT"
    Case adTinyInt: FieldTypeAccess = "BYTE"
    Case adUnsignedTinyInt: FieldTypeAccess = "BYTE"
    Case adSmallInt: FieldTypeAccess = "SHORT"
    Case adInteger: FieldTypeAccess = "LONG"
    Case adSingle: FieldTypeAccess = "SINGLE"
    Case adDouble: FieldTypeAccess = "DOUBLE"
    Case adCurrency: FieldTypeAccess = "CURRENCY"
    Case adDecimal: FieldTypeAccess = "DECIMAL"
    Case adNumeric: FieldTypeAccess = "INTEGER"
    Case adGUID: FieldTypeAccess = "GUID"
    Case adLongVarWChar: FieldTypeAccess = "LONGTEXT"
    Case adVarWChar: FieldTypeAccess = "TEXT"
    Case adDate: FieldTypeAccess = "DATETIME"
    Case adDBTimeStamp: FieldTypeAccess = "DATETIME"
    Case Else
'        MsgBox IntType
         FieldTypeAccess = "-x-"
  End Select
End Function

Public Function AnchoTipoCampo(AdoTipo As ADODB.Field) As String
Dim ch As String
  With AdoTipo
  'MsgBox .Name
  Select Case .Type
    Case TadBoolean:  ch = "Yes"
    Case TadDate, TadDate1: ch = CadDate
    Case TadTime:     ch = CadTime
    Case TadByte:     ch = CadByte
    Case TadInteger:  ch = CadInteger
    Case TadLong:     ch = CadLong
    Case TadSingle:   ch = CadSingle
    Case TadDouble:   ch = CadDouble
    Case TadCurrency: ch = CadCurrency
    Case TadText:     ch = String$(.DefinedSize, "H") & " "
    Case Else:        ch = String$(15, "H") & " "
  End Select
  If ch = "" Then ch = " "
  End With
  AnchoTipoCampo = ch
End Function

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
      Case TadText:     ch = String$(.DefinedSize, "H") & " "
      Case Else:        ch = String$(15, "H") & " "
    End Select
  End With
If ch = "" Then ch = Ninguno
AnchoTipoCampoTexto = Len(ch)
End Function

Public Sub SetFields(DtaTemp As Adodc, _
                     NombCampo As String, _
                     Valor As Variant)
  NombCampo = TrimStrg(NombCampo)
  If NombCampo <> "ID" Then
    With DtaTemp.Recordset
      If IsNull(Valor) Or IsEmpty(Valor) Then
         Select Case .fields(NombCampo).Type
           Case TadBoolean
                Valor = False
           Case TadDate, TadDate1
                Valor = FechaSistema
           Case TadTime
                Valor = Time
           Case TadByte, TadInteger, TadLong, TadCurrency, TadSingle, TadDouble
                Valor = 0
           Case TadText, TadMemo
                Valor = Ninguno
           Case Else
                Valor = Ninguno
         End Select
      End If
       
      Select Case .fields(NombCampo).Type
        Case TadText, TadMemo
             If Len(Valor) > .fields(NombCampo).DefinedSize Then
                Valor = TrimStrg(MidStrg(Valor, 1, .fields(NombCampo).DefinedSize))
             End If
             If Len(Valor) = 0 Then Valor = Ninguno
        Case TadDate, TadDate1
             If Valor = 0 Then Valor = FechaSistema
             If Valor = "00/00/0000" Then Valor = FechaSistema
             If Valor = Ninguno Then Valor = FechaSistema
        Case TadByte, TadInteger, TadLong, TadCurrency, TadSingle, TadDouble
             If Valor > 9999999999.99 Then Valor = 0
      End Select
      'MsgBox .Fields(NombCampo) & vbCrLf & Valor
     .fields(NombCampo) = Valor
    End With
  End If
End Sub

Public Sub SetAddNew(DtaTemp As Adodc)
  DtaTemp.Recordset.AddNew
End Sub

Public Sub SetUpdate(DtaTemp As Adodc)
Dim IUp As Integer
 'MsgBox DtaTemp.RecordSource
  With DtaTemp.Recordset
    For IUp = 0 To .fields.Count - 1
      If .fields(IUp).Name <> "ID" Then
         If IsNull(.fields(IUp)) Or IsEmpty(.fields(IUp)) Then
            Select Case .fields(IUp).Type
              Case TadBoolean
                  .fields(IUp) = False
              Case TadDate, TadDate1
                  .fields(IUp) = FechaSistema
              Case TadTime
                  .fields(IUp) = Time
              Case TadByte, TadInteger, TadCurrency, TadSingle, TadDouble
                  .fields(IUp) = 0
              Case TadLong
                  .fields(IUp) = 0
              Case TadText, TadMemo
                  .fields(IUp) = Ninguno
              Case Else
                  .fields(IUp) = Ninguno
            End Select
         End If
         Select Case .fields(IUp).Type
           Case TadBoolean
               '.Fields(IUp) = False
           Case TadDate, TadDate1
                If CFechaLong(.fields(IUp)) <= CFechaLong("01/01/1900") Then .fields(IUp) = "01/01/1900"
           Case TadTime
               '.Fields(IUp) = Time
           Case TadByte, TadInteger, TadLong, TadCurrency, TadSingle, TadDouble
               '.Fields(IUp) = 0
           Case TadText, TadMemo
               '.Fields(IUp) = Ninguno
           Case Else
               '.Fields(IUp) = Ninguno
         End Select
         If .fields(IUp).Name = "Item" Then
             If IsNull(.fields(IUp)) Or IsEmpty(.fields(IUp)) Then .fields(IUp) = NumEmpresa
             If Len(.fields(IUp)) < 3 Then .fields(IUp) = NumEmpresa
         End If
        'MsgBox .Fields(IUp).Name & vbCrLf & .Fields(IUp)
      End If
    Next IUp
   .Update
  End With
 'MsgBox DtaTemp.RecordSource
End Sub

Public Sub SetAdoAddNew(NombreTabla As String, Optional SinElItem As Boolean)
Dim RegAdodc As ADODB.Recordset
Dim DatosSelect As String
Dim IndDato As Long
  RatonReloj
  NombreTabla = TrimStrg(NombreTabla)
  DatosSelect = "SELECT TOP 0 * " _
              & "FROM " & NombreTabla & " " _
              & " "
''' ****
'''  If NombreTabla = "Clientes" Then
'''     DatosSelect = DatosSelect & "WHERE Grupo = '.' "
'''  Else
'''     If SinElItem = False Then DatosSelect = DatosSelect & "WHERE Item = '.' "
'''  End If
  DatosSelect = CompilarSQL(DatosSelect)
 'MsgBox DatosSelect
  Set RegAdodc = New ADODB.Recordset
  RegAdodc.CursorType = adOpenStatic
  RegAdodc.CursorLocation = adUseClient
  RegAdodc.open DatosSelect, AdoStrCnn, , , adCmdText
  With RegAdodc
       ReDim DatosTabla(.fields.Count + 1) As Campos_Tabla
       DatosTabla(0).Campo = NombreTabla
       DatosTabla(0).Ancho = .fields.Count
       DatosTabla(0).Valor = 0
       DatosTabla(0).Tipo = 0
       For IndDato = 0 To .fields.Count - 1
           DatosTabla(IndDato + 1).Campo = .fields(IndDato).Name
           DatosTabla(IndDato + 1).Ancho = .fields(IndDato).DefinedSize
           DatosTabla(IndDato + 1).Tipo = .fields(IndDato).Type
           Select Case .fields(IndDato).Type
             Case TadBoolean
                  DatosTabla(IndDato + 1).Valor = adFalse
             Case TadText, TadMemo
                  DatosTabla(IndDato + 1).Valor = Ninguno
             Case TadDate, TadDate1
                  DatosTabla(IndDato + 1).Valor = FechaSistema
             Case TadByte, TadInteger, TadLong, TadCurrency, TadSingle, TadDouble
                  DatosTabla(IndDato + 1).Valor = 0
           End Select
       Next IndDato
  End With
  RatonNormal
End Sub

Public Sub SetAdoFields(NombCampo As String, ValorCampo As Variant)
Dim IndDato As Long
  RatonReloj
  NombCampo = TrimStrg(NombCampo)
  If IsNull(ValorCampo) Or IsEmpty(ValorCampo) Then ValorCampo = Null
  For IndDato = 1 To DatosTabla(0).Ancho
   If DatosTabla(IndDato).Campo = NombCampo Then
      Select Case DatosTabla(IndDato).Tipo
        Case TadText, TadMemo
             If IsNull(ValorCampo) Or IsEmpty(ValorCampo) Then ValorCampo = ""
             ValorCampo = Replace(ValorCampo, "'", "`")
             ValorCampo = Replace(ValorCampo, "#", "No.")
             If Len(ValorCampo) > DatosTabla(IndDato).Ancho Then
                ValorCampo = TrimStrg(MidStrg(ValorCampo, 1, DatosTabla(IndDato).Ancho))
             End If
             If Len(ValorCampo) = 0 Then ValorCampo = Ninguno
        Case TadDate, TadDate1
             If ValorCampo = 0 Or ValorCampo = Ninguno Then ValorCampo = FechaSistema
             If IsNull(ValorCampo) Then ValorCampo = FechaSistema
             If Not IsDate(ValorCampo) Then ValorCampo = FechaSistema
        Case TadByte: If ValorCampo > 255 Then ValorCampo = 0
        Case TadInteger: If ValorCampo > 32767 Then ValorCampo = 0
        Case TadLong: If ValorCampo > 2147483647 Then ValorCampo = 0
        Case TadCurrency: If ValorCampo > 999999999999.99 Then ValorCampo = 0
        Case TadSingle: If ValorCampo > 9999999999.99 Then ValorCampo = 0
        Case TadDouble: If ValorCampo > 999999999999.99 Then ValorCampo = 0
      End Select
     'MsgBox "SetAdoFields - " & NombCampo & " = " & ValorCampo
      DatosTabla(IndDato).Valor = ValorCampo
   End If
  Next IndDato
  RatonNormal
End Sub

Public Sub SetAdoUpdate()
Dim AdoCon1 As ADODB.Connection
Dim DatosSelect As String
Dim InsertarCampos As String
Dim InsertDato As Variant
Dim IndDato As Long
Dim IdTime As Long
  RatonReloj
  InsertarCampos = ""
  DatosSelect = "INSERT INTO " & DatosTabla(0).Campo & " ("
  For IndDato = 1 To DatosTabla(0).Ancho
      If DatosTabla(IndDato).Campo <> "ID" Then DatosSelect = DatosSelect & "[" & DatosTabla(IndDato).Campo & "],"
  Next IndDato
  If MidStrg(DatosSelect, Len(DatosSelect), 1) = "," Then DatosSelect = MidStrg(DatosSelect, 1, Len(DatosSelect) - 1)
  DatosSelect = DatosSelect & ") VALUES ("
  For IndDato = 1 To DatosTabla(0).Ancho
     'MsgBox DatosTabla(IndDato).Campo & " = " & DatosTabla(IndDato).Valor
     'MsgBox DatosSelect
      If DatosTabla(IndDato).Campo <> "ID" Then
        Select Case DatosTabla(IndDato).Tipo
          Case TadBoolean
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               If DatosTabla(IndDato).Valor = Ninguno Or DatosTabla(IndDato).Valor = "" Then DatosTabla(IndDato).Valor = 0
               If DatosTabla(IndDato).Valor < 0 Then DatosTabla(IndDato).Valor = 1
               If DatosTabla(IndDato).Valor > 1 Then DatosTabla(IndDato).Valor = 1
              'MsgBox DatosTabla(IndDato).Valor
               DatosSelect = DatosSelect & CByte(DatosTabla(IndDato).Valor)
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CByte(DatosTabla(IndDato).Valor) & vbCrLf
          Case TadText, TadMemo
               If IsNull(DatosTabla(IndDato).Valor) Or IsEmpty(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = Ninguno
               If DatosTabla(IndDato).Campo = "Periodo" And DatosTabla(IndDato).Valor = Ninguno Then
                  DatosTabla(IndDato).Valor = Periodo_Contable
               End If
               If DatosTabla(IndDato).Campo = "Item" And DatosTabla(IndDato).Valor = Ninguno Then
                  DatosTabla(IndDato).Valor = NumEmpresa
               End If
               If DatosTabla(IndDato).Campo = "CodigoU" And DatosTabla(IndDato).Valor = Ninguno Then
                  DatosTabla(IndDato).Valor = CodigoUsuario
               End If
               DatosSelect = DatosSelect & "'" & DatosTabla(IndDato).Valor & "'"
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & "'" & DatosTabla(IndDato).Valor & "'" & vbCrLf
          Case TadDate, TadDate1
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = FechaSistema
               DatosSelect = DatosSelect & "#" & BuscarFecha(CStr(DatosTabla(IndDato).Valor)) & "#"
              'MsgBox DatosSelect & vbCrLf & DatosTabla(IndDato).Valor
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & "#" & BuscarFecha(CStr(DatosTabla(IndDato).Valor)) & "#" & vbCrLf
          Case TadByte
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CByte(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CByte(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadInteger
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CInt(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CInt(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadLong
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CLng(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CLng(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadCurrency
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CCur(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CCur(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadSingle
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CSng(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CSng(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadDouble
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CDbl(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CSng(Val(DatosTabla(IndDato).Valor)) & vbCrLf
        End Select
        DatosSelect = DatosSelect & ","
      End If
  Next IndDato
  DatosSelect = DatosSelect & ");"
  DatosSelect = Replace(DatosSelect, ",)", ")")
  DatosSelect = CompilarSQL(DatosSelect)
 'MsgBox DatosSelect
 'MsgBox InsertarCampos
  
  Ejecutar_SQL_SP DatosSelect
'''  Set AdoCon1 = New ADODB.Connection
'''  AdoCon1.open AdoStrCnn
'''  AdoCon1.Execute DatosSelect, RegAfectados, adCmdText
  RatonNormal
End Sub

Public Sub SetAdoUpdateMySQL(AdoStrCnnMySQL As String)
Dim AdoCon1 As ADODB.Connection
Dim DatosSelect As String
Dim InsertarCampos As String
Dim InsertDato As Variant
Dim IndDato As Long
Dim IdTime As Long
Dim MySQLServer As Boolean
  RatonReloj
  MySQLServer = SQL_Server
  SQL_Server = False
  InsertarCampos = ""
  DatosSelect = "INSERT INTO " & DatosTabla(0).Campo & " ("
  For IndDato = 1 To DatosTabla(0).Ancho
      If DatosTabla(IndDato).Campo <> "ID" Then DatosSelect = DatosSelect & "`" & DatosTabla(IndDato).Campo & "`,"
  Next IndDato
  If MidStrg(DatosSelect, Len(DatosSelect), 1) = "," Then DatosSelect = MidStrg(DatosSelect, 1, Len(DatosSelect) - 1)
  DatosSelect = DatosSelect & ") VALUES ("
  For IndDato = 1 To DatosTabla(0).Ancho
     'MsgBox DatosTabla(IndDato).Campo & " = " & DatosTabla(IndDato).Valor
     'MsgBox DatosSelect
      If DatosTabla(IndDato).Campo <> "ID" Then
        Select Case DatosTabla(IndDato).Tipo
          Case TadBoolean
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               If DatosTabla(IndDato).Valor = Ninguno Or DatosTabla(IndDato).Valor = "" Then DatosTabla(IndDato).Valor = 0
               If DatosTabla(IndDato).Valor < 0 Then DatosTabla(IndDato).Valor = 1
               If DatosTabla(IndDato).Valor > 1 Then DatosTabla(IndDato).Valor = 1
              'MsgBox DatosTabla(IndDato).Valor
               DatosSelect = DatosSelect & CByte(DatosTabla(IndDato).Valor)
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CByte(DatosTabla(IndDato).Valor) & vbCrLf
          Case TadText, TadMemo
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = Ninguno
               If DatosTabla(IndDato).Campo = "Periodo" And DatosTabla(IndDato).Valor = Ninguno Then
                  DatosTabla(IndDato).Valor = Periodo_Contable
               End If
               If DatosTabla(IndDato).Campo = "Item" And DatosTabla(IndDato).Valor = Ninguno Then
                  DatosTabla(IndDato).Valor = NumEmpresa
               End If
               If DatosTabla(IndDato).Campo = "CodigoU" And DatosTabla(IndDato).Valor = Ninguno Then
                  DatosTabla(IndDato).Valor = CodigoUsuario
               End If
               DatosSelect = DatosSelect & "'" & DatosTabla(IndDato).Valor & "'"
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & "'" & DatosTabla(IndDato).Valor & "'" & vbCrLf
          Case TadDate, TadDate1
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = FechaSistema
               DatosSelect = DatosSelect & "'" & BuscarFecha(CStr(DatosTabla(IndDato).Valor)) & "'"
              'MsgBox DatosSelect & vbCrLf & DatosTabla(IndDato).Valor
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & "'" & BuscarFecha(CStr(DatosTabla(IndDato).Valor)) & "'" & vbCrLf
          Case TadByte
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CByte(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CByte(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadInteger
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CInt(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CInt(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadLong
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CLng(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CLng(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadCurrency
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CCur(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CCur(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadSingle
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CSng(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CSng(Val(DatosTabla(IndDato).Valor)) & vbCrLf
          Case TadDouble
               If IsNull(DatosTabla(IndDato).Valor) Then DatosTabla(IndDato).Valor = 0
               DatosSelect = DatosSelect & CDbl(Val(DatosTabla(IndDato).Valor))
               InsertarCampos = InsertarCampos & DatosTabla(IndDato).Campo & " = " & CSng(Val(DatosTabla(IndDato).Valor)) & vbCrLf
        End Select
        DatosSelect = DatosSelect & ","
      End If
  Next IndDato
  DatosSelect = DatosSelect & ");"
  DatosSelect = Replace(DatosSelect, ",)", ")")
  DatosSelect = CompilarSQL(DatosSelect)
 'MsgBox AdoStrCnnMySQL & vbCrLf & vbCrLf & DatosSelect
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnnMySQL
  AdoCon1.Execute DatosSelect, RegAfectados, adCmdText
  SQL_Server = MySQLServer
  RatonNormal
End Sub

Public Sub CopiarAdoTabla(NombreTabla As String, _
                          NewItemEmpresa As String)
Dim AdoCon1 As ADODB.Connection
Dim RegAdodc As ADODB.Recordset

Dim DatosSelect As String
Dim InsertDato As Variant
Dim IndDato As Long
Dim IdTime As Long
  RatonReloj
  Si_No = False
  NombreTabla = TrimStrg(NombreTabla)
  DatosSelect = "SELECT * " _
              & "FROM " & NombreTabla & " " _
              & "WHERE Item = '.' "
  DatosSelect = CompilarSQL(DatosSelect)
  Set RegAdodc = New ADODB.Recordset
  RegAdodc.CursorType = adOpenStatic
  RegAdodc.CursorLocation = adUseClient
  RegAdodc.open DatosSelect, AdoStrCnn, , , adCmdText
  With RegAdodc
       ReDim DatosTabla(.fields.Count + 1) As Campos_Tabla
       DatosTabla(0).Campo = NombreTabla
       DatosTabla(0).Ancho = .fields.Count
       DatosTabla(0).Valor = 0
       DatosTabla(0).Tipo = 0
       For IndDato = 0 To .fields.Count - 1
           DatosTabla(IndDato + 1).Campo = .fields(IndDato).Name
           DatosTabla(IndDato + 1).Ancho = 0
           DatosTabla(IndDato + 1).Tipo = 0
           DatosTabla(IndDato + 1).Valor = 0
           If DatosTabla(IndDato + 1).Campo = "Periodo" Then Si_No = True
       Next IndDato
  End With
''NewItemEmpresa
  DatosSelect = "DELETE * " _
              & "FROM " & NombreTabla & " " _
              & "WHERE Item = '" & NewItemEmpresa & "' "
  If Si_No Then DatosSelect = DatosSelect & "AND Periodo = '" & Periodo_Contable & "' "
  Ejecutar_SQL_SP DatosSelect
  DatosSelect = CompilarSQL(DatosSelect)
  'Borramos datos si existen en la empresa nueva
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  AdoCon1.Execute DatosSelect, RegAfectados, adCmdText
 'Insertamos toda la Tabla
  DatosSelect = "INSERT INTO " & NombreTabla & " ("
  For IndDato = 1 To DatosTabla(0).Ancho
      If IndDato = DatosTabla(0).Ancho Then
         DatosSelect = DatosSelect & DatosTabla(IndDato).Campo & ") "
      Else
         DatosSelect = DatosSelect & DatosTabla(IndDato).Campo & ", "
      End If
  Next IndDato
  DatosSelect = DatosSelect & "SELECT "
  For IndDato = 1 To DatosTabla(0).Ancho
      If DatosTabla(IndDato).Campo = "Item" Then
         CodigoB = "'" & NewItemEmpresa & "' As XItem"
      Else
         CodigoB = DatosTabla(IndDato).Campo
      End If
      DatosSelect = DatosSelect & CodigoB
      If IndDato = DatosTabla(0).Ancho Then
         DatosSelect = DatosSelect & " "
      Else
         DatosSelect = DatosSelect & ", "
      End If
  Next IndDato
  DatosSelect = DatosSelect & "FROM " & NombreTabla & " " _
              & "WHERE Item = '" & NumEmpresa & "' "
  If Si_No Then DatosSelect = DatosSelect & "AND Periodo = '" & Periodo_Contable & "' "
  DatosSelect = CompilarSQL(DatosSelect)
  'MsgBox DatosSelect
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  AdoCon1.Execute DatosSelect, RegAfectados, adCmdText
  RatonNormal
End Sub

Public Sub Copiar_Tabla_Empresa(NombreTabla As String, _
                                OldItemEmpresa As String, _
                                PeriodoCopy As String, _
                                Optional AdoStrCnnCopy As String, _
                                Optional NoBorrarTabla As Boolean)

Dim AdoCon1   As ADODB.Connection
Dim RegAdodc  As ADODB.Recordset
Dim CopyAdodc As ADODB.Recordset

Dim AdoStrCnnTemp As String
Dim DatosSelect   As String
Dim InsertDato    As Variant
Dim InsFields     As Boolean
Dim IndDato       As Long
Dim IdTime        As Long

  RatonReloj
  NombreTabla = TrimStrg(NombreTabla)
  Progreso_Barra.Mensaje_Box = "[00%] Copiando de " & NombreTabla
  Progreso_Esperar True
 '& "WHERE Item = '" & NumEmpresa & "' "
  DatosSelect = "SELECT * " _
              & "FROM " & NombreTabla & " " _
              & "WHERE 1 = 0 "
  Select_AdoDB RegAdodc, DatosSelect
  With RegAdodc
       ReDim DatosTabla(.fields.Count + 1) As Campos_Tabla
       DatosTabla(0).Campo = NombreTabla
       DatosTabla(0).Ancho = .fields.Count
       DatosTabla(0).Valor = 0
       DatosTabla(0).Tipo = 0
       DatosTabla(0).Si_Periodo = False
       For IndDato = 0 To .fields.Count - 1
           DatosTabla(IndDato + 1).Campo = .fields(IndDato).Name
           DatosTabla(IndDato + 1).Ancho = 0
           DatosTabla(IndDato + 1).Tipo = 0
           DatosTabla(IndDato + 1).Valor = 0
           If DatosTabla(IndDato + 1).Campo = "Periodo" Then DatosTabla(0).Si_Periodo = True
       Next IndDato
  End With
  RegAdodc.Close

 'Borramos datos si existen en la empresa nueva
  If Not NoBorrarTabla Then
     DatosSelect = "DELETE * " _
                 & "FROM " & NombreTabla & " " _
                 & "WHERE Item = '" & NumEmpresa & "' "
     If DatosTabla(0).Si_Periodo Then
        If IsDate(PeriodoCopy) Then
           DatosSelect = DatosSelect & "AND Periodo = '" & PeriodoCopy & "' "
        Else
           DatosSelect = DatosSelect & "AND Periodo = '" & Periodo_Contable & "' "
        End If
     End If
    'MsgBox DatosSelect
     Ejecutar_SQL_SP DatosSelect
  End If
 'Insertamos toda la Tabla de registro en registro
  Contador = 0
  If Len(AdoStrCnnCopy) <= 1 Then AdoStrCnnCopy = AdoStrCnn
  AdoStrCnnTemp = AdoStrCnn
  AdoStrCnn = AdoStrCnnCopy
  DatosSelect = "SELECT * " _
              & "FROM " & NombreTabla & " " _
              & "WHERE Item = '" & OldItemEmpresa & "' "
  If DatosTabla(0).Si_Periodo Then
     If IsDate(PeriodoCopy) Then
        DatosSelect = DatosSelect & "AND Periodo = '" & PeriodoCopy & "' "
     Else
        DatosSelect = DatosSelect & "AND Periodo = '" & Periodo_Contable & "' "
     End If
  End If
  Select_AdoDB CopyAdodc, DatosSelect
  AdoStrCnn = AdoStrCnnTemp
  With CopyAdodc
   If .RecordCount > 0 Then
      'MsgBox AdoStrCnnTemp & vbCrLf & vbCrLf & NombreTabla & vbCrLf & vbCrLf & .RecordCount
       Do While Not .EOF
          Progreso_Barra.Mensaje_Box = "[" & Format(Contador / .RecordCount, "00%") & "] Copiando de " & NombreTabla
          Progreso_Esperar True
          SetAdoAddNew NombreTabla
          For IndDato = 0 To .fields.Count - 1
              InsFields = True
              Select Case .fields(IndDato).Name
                Case "ID": InsFields = False
                Case "Item": InsFields = False
              End Select
              If InsFields Then SetAdoFields .fields(IndDato).Name, .fields(IndDato)
          Next IndDato
          SetAdoUpdate
          Contador = Contador + 1
         .MoveNext
       Loop
   End If
  End With
  CopyAdodc.Close
   
'''  DatosSelect = "INSERT INTO " & NombreTabla & " ("
'''  For IndDato = 1 To DatosTabla(0).Ancho
'''      If DatosTabla(IndDato).Campo <> "ID" Then DatosSelect = DatosSelect & DatosTabla(IndDato).Campo & ", "
'''  Next IndDato
'''  DatosSelect = TrimStrg(DatosSelect)
'''  DatosSelect = MidStrg(DatosSelect, 1, Len(DatosSelect) - 1) & ") "
'''  DatosSelect = DatosSelect & "SELECT "
'''  For IndDato = 1 To DatosTabla(0).Ancho
'''      If DatosTabla(IndDato).Campo = "Item" Then
'''         CodigoB = "'" & NumEmpresa & "' As XItem"
'''      ElseIf DatosTabla(IndDato).Campo = "Periodo" Then
'''         If IsDate(PeriodoCopy) Then
'''            CodigoB = "'" & PeriodoCopy & "' As XPeriodo"
'''         Else
'''            CodigoB = "'" & Periodo_Contable & "' As XPeriodo"
'''         End If
'''      Else
'''         CodigoB = DatosTabla(IndDato).Campo
'''      End If
'''      If CodigoB <> "ID" Then DatosSelect = DatosSelect & CodigoB & ", "
'''  Next IndDato
'''  DatosSelect = TrimStrg(DatosSelect)
'''  DatosSelect = MidStrg(DatosSelect, 1, Len(DatosSelect) - 1) & " "
'''  DatosSelect = DatosSelect & "FROM " & NombreTabla & " " _
'''              & "WHERE Item = '" & OldItemEmpresa & "' "
'''  If DatosTabla(0).Si_Periodo Then
'''     If IsDate(PeriodoCopy) Then
'''        DatosSelect = DatosSelect & "AND Periodo = '" & PeriodoCopy & "' "
'''     Else
'''        DatosSelect = DatosSelect & "AND Periodo = '" & Periodo_Contable & "' "
'''     End If
'''  End If
''' 'MsgBox DatosSelect
'''  Ejecutar_SQL_SP DatosSelect
  Progreso_Barra.Mensaje_Box = "Ok"
  Progreso_Esperar True
  RatonNormal
End Sub

Public Sub CopiarAdoTablaPeriodo(NombreTabla As String, _
                                 NewPeriodoEmpresa As String)
Dim AdoCon1 As ADODB.Connection
Dim RegAdodc As ADODB.Recordset

Dim DatosSelect As String
Dim InsertDato As Variant
Dim IndDato As Long
Dim IdTime As Long
  Si_No = False
  RatonReloj
  NombreTabla = TrimStrg(NombreTabla)
  DatosSelect = "SELECT * " _
              & "FROM " & NombreTabla & " " _
              & "WHERE Item = '.' "
  DatosSelect = CompilarSQL(DatosSelect)
  Set RegAdodc = New ADODB.Recordset
  RegAdodc.CursorType = adOpenStatic
  RegAdodc.CursorLocation = adUseClient
  RegAdodc.open DatosSelect, AdoStrCnn, , , adCmdText
  With RegAdodc
       ReDim DatosTabla(.fields.Count + 1) As Campos_Tabla
       DatosTabla(0).Campo = NombreTabla
       DatosTabla(0).Ancho = .fields.Count
       DatosTabla(0).Valor = 0
       DatosTabla(0).Tipo = 0
       For IndDato = 0 To .fields.Count - 1
           DatosTabla(IndDato + 1).Campo = .fields(IndDato).Name
           DatosTabla(IndDato + 1).Ancho = 0
           DatosTabla(IndDato + 1).Tipo = 0
           DatosTabla(IndDato + 1).Valor = 0
           If DatosTabla(IndDato + 1).Campo = "Periodo" Then Si_No = True
       Next IndDato
  End With
''NewPeriodoEmpresa
  DatosSelect = "DELETE * " _
              & "FROM " & NombreTabla & " " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & NewPeriodoEmpresa & "' "
  Ejecutar_SQL_SP DatosSelect
  DatosSelect = CompilarSQL(DatosSelect)
  
  'Borramos datos si existen en la empresa nueva
  'MsgBox DatosSelect
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  AdoCon1.Execute DatosSelect, RegAfectados, adCmdText
 'Insertamos toda la Tabla
  DatosSelect = "INSERT INTO " & NombreTabla & " ("
  For IndDato = 1 To DatosTabla(0).Ancho
      If DatosTabla(IndDato).Campo <> "ID" Then
         DatosSelect = DatosSelect & DatosTabla(IndDato).Campo & ","
      End If
  Next IndDato
  DatosSelect = MidStrg(DatosSelect, 1, Len(DatosSelect) - 1)
  DatosSelect = DatosSelect & ") "
  DatosSelect = DatosSelect _
              & "SELECT "
  For IndDato = 1 To DatosTabla(0).Ancho
      If DatosTabla(IndDato).Campo = "Periodo" Then
         CodigoB = "'" & NewPeriodoEmpresa & "' As XPeriodo"
      Else
         CodigoB = DatosTabla(IndDato).Campo
      End If
      If CodigoB <> "ID" Then DatosSelect = DatosSelect & CodigoB & ","
  Next IndDato
  DatosSelect = MidStrg(DatosSelect, 1, Len(DatosSelect) - 1)
  DatosSelect = DatosSelect & " " _
              & "FROM " & NombreTabla & " " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' "
  DatosSelect = CompilarSQL(DatosSelect)
  'MsgBox DatosSelect
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  AdoCon1.Execute DatosSelect, RegAfectados, adCmdText
  RatonNormal
End Sub

Public Sub CopiarRegistro(DtaTemp As Adodc, _
                          Campo As String, _
                          Valor As Variant, _
                          NumEmp2 As String)
Dim IUp As Integer
Dim NombreCampo As String
Dim ValorCampo As Variant
  With DtaTemp.Recordset
   If .RecordCount > 0 Then
       ReDim TipoC(.fields.Count) As Campos_Tabla
       For IUp = 0 To .fields.Count - 1
           TipoC(IUp).Campo = .fields(IUp).Name
           TipoC(IUp).Valor = .fields(IUp)
       Next IUp
       SetAddNew DtaTemp
       For IUp = 0 To .fields.Count - 1
           NombreCampo = .fields(IUp).Name
           ValorCampo = .fields(IUp)
           SetFields DtaTemp, TipoC(IUp).Campo, TipoC(IUp).Valor
       Next IUp
       SetFields DtaTemp, Campo, Valor
       SetFields DtaTemp, "Item", NumEmp2
       SetFields DtaTemp, "Grupo", NumEmp2
       SetUpdate DtaTemp
   End If
  End With
End Sub

Public Function ReadField(DtaTemp As Adodc, _
                          NombCampo As String) As Variant
Dim Valor As Variant
  With DtaTemp.Recordset
    If IsNull(.fields(NombCampo)) Then
       Select Case .fields(NombCampo).Type
         Case TadBoolean
              Valor = False
         Case TadDate, TadDate1
              Valor = FechaSistema
         Case TadTime
              Valor = Time
         Case TadByte, TadInteger, TadLong, TadCurrency, TadSingle, TadDouble
              Valor = 0
         Case TadText, TadMemo
              Valor = Ninguno
         Case Else
              Valor = Ninguno
       End Select
    Else
       Valor = .fields(NombCampo)
    End If
  End With
  ReadField = Valor
End Function

Public Function CampoWidth(AdoTipo As ADODB.Field) As Single
Dim DistCampo As Single
Dim ch As String
DistCampo = 0
StrgAncho = AnchoTipoCampo(AdoTipo)
AltoLetra = Printer.TextHeight(StrgAncho)
LongNumero = Printer.TextWidth(StrgAncho)

StrgFormatoCampo = FormatoTipoCampo(AdoTipo)
'MsgBox StrgFormatoCampo
LongCampo = Printer.TextWidth(StrgFormatoCampo)
Select Case AdoTipo.Type
  Case TadByte, TadCurrency, TadInteger, TadLong, TadSingle, TadDouble
       DistCampo = LongNumero - LongCampo
       If DistCampo <= 0 Then DistCampo = 0.05
End Select
If DistCampo <= 0 Then DistCampo = 0.05
CampoWidth = DistCampo
End Function

Public Sub PrinterAllFields(CantidadCampos As Integer, _
                            Yo As Single, _
                            DataAllFields As Adodc, _
                            PonerLineas As Boolean, _
                            NombresCampos As Boolean, _
                            Optional SegundaPagina As Boolean)
Dim Xi As Single
Dim ItemCampo As Integer
Dim CInicio As Integer
Dim CFinal As Integer
'If Yo < LimiteAlto Then
AltoLetra = Printer.TextHeight("H")
CFinal = CantidadCampos - 1

With DataAllFields.Recordset
  If NombresCampos Then
     Printer.FontBold = True
     Printer.FontItalic = True
     Printer.FontUnderline = True
  End If
  Xi = 0: LimpiarLinea Xi, Yo
  If SegundaPagina Then
     CInicio = EnDosPaginas + 1
     CFinal = CantidadCampos - 1
  Else
     CInicio = 0
     If EnDosPaginas > 0 Then CFinal = EnDosPaginas Else CFinal = CantidadCampos - 1
  End If
  If CInicio > CFinal Then CInicio = CFinal
  'MsgBox CInicio & vbCrLf & CFinal
  For ItemCampo = CInicio To CFinal
      Xi = Ancho(ItemCampo)
      If ((Xi > 0) And (Yo > 0)) Then
         If Xi < AnchoPapel Then
            'MsgBox Xi & vbCrLf & Yo & vbCrLf & PonerLineas
            If NombresCampos Then
               Distancia = 0.05
               StrgFormatoCampo = .fields(ItemCampo).Name
            Else
               LimpiarLinea Xi, Yo, PonerLineas
               Distancia = CampoWidth(.fields(ItemCampo))
            End If
           ' If Yo <= LimiteAlto Then
               PictPrint_Cuadro_Linea Printer, Xi + 0.05, Yo, AnchoPapel, Yo + 0.4, Blanco, "BF"
               Printer.CurrentX = Xi + Distancia
               Printer.CurrentY = Yo
               If StrgFormatoCampo = Ninguno Then StrgFormatoCampo = " "
               If StrgFormatoCampo = "0" Then StrgFormatoCampo = " "
               If StrgFormatoCampo = "0.00" Then StrgFormatoCampo = " "
               Printer.Print StrgFormatoCampo
              'MsgBox StrgFormatoCampo
           ' End If
         End If
      End If
  Next ItemCampo
  If PonerLineas Then
     LongNumero = Printer.TextWidth(AnchoTipoCampo(.fields(CFinal))) + 0.1
     If Ancho(CFinal) > 0 Then Imprimir_Linea_V Ancho(CFinal) + LongNumero, Yo, Yo + 0.35
  End If
End With
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontUnderline = False
'End If
End Sub

Public Sub PrinterFieldText(Xo As Single, _
                            Yo As Single, _
                            Texto As String, _
                            AdoTipo As ADODB.Field, _
                            Optional ImpLineaCero As Boolean)
Dim CantStrg As Single
'If Yo < LimiteAlto Then
If ((Xo > 0) And (Yo > 0)) Then
   CantStrg = 0
   If Texto <> "" Then CantStrg = Printer.TextWidth(Texto) + 0.1
   Distancia = CampoWidth(AdoTipo)
   If Yo <= LimiteAlto Then
      Printer.CurrentX = Xo
      Printer.CurrentY = Yo
      If Texto = Ninguno Then Texto = " "
      If Texto = "0" Then Texto = " "
      Printer.Print Texto
      Printer.CurrentX = Xo + Distancia + CantStrg
      Printer.CurrentY = Yo
      If StrgFormatoCampo = Ninguno Then
         StrgFormatoCampo = " "
      ElseIf StrgFormatoCampo = "0" Or StrgFormatoCampo = "0.00" Then
         If ImpLineaCero Then StrgFormatoCampo = "--" Else StrgFormatoCampo = " "
      End If
      Printer.Print StrgFormatoCampo
   End If
End If
'End If
End Sub

Public Sub Ancho_Recordset(IniciarX As Single, _
                           Datas As ADODB.Recordset, _
                           PorteLetra As Integer, _
                           TipoLetra As String, _
                           Orientacion As Byte, _
                           Optional EsCampoCorto As Boolean)
Dim IniXX As Single
Dim SumaTotalAncho As Single
Dim AnchoCampo As Single
EnDosPaginas = 0
Printer.Orientation = Orientacion_Pagina      'Orientacion_Pagina normal de la hoja
Printer.ScaleMode = vbCentimeters   'Escala de centimetros
If SetPapelPRN > 0 Then Printer.PaperSize = SetPapelPRN Else Printer.PaperSize = vbPRPSLetter
Printer.DrawWidth = 1               'Ancho de la línea
LimiteAlto = Printer.ScaleHeight - 1.5 'Limite de impresión a lo largo
LimiteAncho = Printer.ScaleWidth       'Limite de impresión a lo ancho
AnchoPapel = Redondear(Printer.ScaleWidth - 0.3, 2) 'Ancho de impresion del papel
Printer.FontName = TipoLetra        'Tipo de letra en todo el sistema
Printer.FontSize = PorteLetra 'Porte de la letra default
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontUnderline = False
SetAnchoCampos Printer, EsCampoCorto
CantCampos = 0
With Datas
     CantCampos = .fields.Count
     ReDim Ancho(CantCampos + 1) As Single
     ReDim AnchoDeCampo(CantCampos + 1) As String
     For I = 0 To CantCampos - 1
         AnchoDeCampo(I) = AnchoTipoCampo(.fields(I))
         Ancho(I) = 0
     Next I
End With
IniXX = IniciarX
If IniXX < 0.5 Then IniXX = 0.5
SumaTotalAncho = IniXX
Ancho(0) = SumaTotalAncho
Evaluar = False
For I = 0 To CantCampos - 1
    AnchoCampo = Printer.TextWidth(AnchoDeCampo(I)) + 0.1
    If AnchoCampo < 0.5 Then AnchoCampo = 0.5
    If AnchoCampo > 6.5 Then AnchoCampo = 6.5
    SumaTotalAncho = SumaTotalAncho + AnchoCampo
    If SumaTotalAncho + 1 >= AnchoPapel Then
       EnDosPaginas = I
       Evaluar = True
    End If
    
    If Evaluar Then
       SumaTotalAncho = IniXX
       Evaluar = False
    End If
    Ancho(I + 1) = SumaTotalAncho
Next I
End Sub

Public Sub DataAnchoCampos(IniciarX As Single, _
                           Datas As Adodc, _
                           PorteLetra As Integer, _
                           TipoLetra As String, _
                           Orientacion As Byte, _
                           Optional EsCampoCorto As Boolean)
Dim IniXX As Single
Dim SumaTotalAncho As Single
Dim AnchoCampo As Single

 EnDosPaginas = 0
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
           LimiteAlto = Redondear(Printer.ScaleHeight - 1.6, 4)
   Case 2: Printer.Orientation = vbPRORLandscape 'Horizontal
           LimiteAlto = Redondear(Printer.ScaleHeight - 1, 4)
   Case Else: Printer.Orientation = vbPRORPortrait
              Orientacion_Pagina = 1
              LimiteAlto = Redondear(Printer.ScaleHeight - 1.6, 4)
 End Select
 Printer.DrawWidth = 1                  'Ancho de la línea
 Printer.FontName = TipoLetra           'Tipo de letra en todo el sistema
 Printer.FontSize = PorteLetra          'Porte de la letra default
 Printer.FontBold = False
 Printer.FontItalic = False
 Printer.FontUnderline = False
 LimiteAncho = Redondear(Printer.ScaleWidth - 0.5, 4)   'Limite de impresión a lo ancho
 SetPapelAncho = Redondear(Printer.ScaleWidth, 4)
 SetPapelLargo = Redondear(Printer.ScaleHeight, 4)
 AnchoPapel = Redondear(Printer.ScaleWidth - 0.3, 4) 'Ancho de impresion del papel
'Determinamos en ancho maximo de cada campo de la tabla para imprimir
 SetAnchoCampos Printer, EsCampoCorto
 IniXX = IniciarX
 If IniXX < 0.5 Then IniXX = 0.5
 SumaTotalAncho = IniXX
 CantCampos = 0
With Datas.Recordset
 If .RecordCount > 0 Then
     CantCampos = .fields.Count
     ReDim Ancho(CantCampos + 1) As Single
     ReDim AnchoDeCampo(CantCampos + 1) As String
     For I = 0 To CantCampos - 1
         AnchoDeCampo(I) = AnchoTipoCampo(.fields(I))
         Ancho(I) = 0
     Next I
     Ancho(0) = SumaTotalAncho
     Evaluar = False
     For I = 0 To CantCampos - 1
         'AnchoCampo = PorteLetra * Len(AnchoDeCampo(I)) * 0.0185567
         AnchoCampo = Printer.TextWidth(AnchoDeCampo(I)) + 0.1
         If AnchoCampo < 0.5 Then AnchoCampo = 0.5
         If AnchoCampo > 6.5 Then AnchoCampo = 6.5
         SumaTotalAncho = SumaTotalAncho + AnchoCampo
         If SumaTotalAncho + 1 >= AnchoPapel Then
            EnDosPaginas = I
            Evaluar = True
         End If
         If Evaluar Then
            SumaTotalAncho = IniXX
            Evaluar = False
         End If
         Ancho(I + 1) = SumaTotalAncho
     Next I
     Ancho(CantCampos) = SumaTotalAncho
 End If
End With
End Sub

Public Sub AdoDBAnchoCampos(IniciarX As Single, _
                           Datas As ADODB.Recordset, _
                           PorteLetra As Integer, _
                           TipoLetra As String, _
                           Orientacion As Byte, _
                           Optional PaginaA4 As Boolean, _
                           Optional EsCampoCorto As Boolean)
Dim IniXX As Single
Dim SumaTotalAncho As Single
Dim AnchoCampo As Single
Printer.Orientation = Orientacion_Pagina      'Orientacion_Pagina normal de la hoja
Printer.ScaleMode = vbCentimeters   'Escala de centimetros
If SetPapelPRN > 0 Then Printer.PaperSize = SetPapelPRN Else Printer.PaperSize = vbPRPSLetter
Printer.DrawWidth = 1               'Ancho de la línea
LimiteAlto = Printer.ScaleHeight - 1.5 'Limite de impresión a lo largo
LimiteAncho = Printer.ScaleWidth       'Limite de impresión a lo largo
AnchoPapel = Printer.ScaleWidth        'Limite de impresión a lo largo
SetPapelAncho = Printer.ScaleWidth
SetPapelLargo = Printer.ScaleHeight
Printer.FontName = TipoLetra        'Tipo de letra en todo el sistema
Printer.FontSize = PorteLetra 'Porte de la letra default
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontUnderline = False

CantCampos = 0
With Datas
     CantCampos = .fields.Count
     ReDim Ancho(CantCampos + 1) As Single
     ReDim AnchoDeCampo(CantCampos + 1) As String
     For I = 0 To CantCampos - 1
         AnchoDeCampo(I) = AnchoTipoCampo(.fields(I))
         Ancho(I) = 0
     Next I
End With
IniXX = IniciarX
If IniXX < 0.5 Then IniXX = 0.5
SumaTotalAncho = IniXX
Ancho(0) = SumaTotalAncho
For I = 0 To CantCampos - 1
    AnchoCampo = Printer.TextWidth(AnchoDeCampo(I)) + 0.1
    If AnchoCampo < 0.35 Then AnchoCampo = 0.35
    If AnchoCampo > 6.5 Then AnchoCampo = 6.5
    SumaTotalAncho = SumaTotalAncho + AnchoCampo
    Ancho(I + 1) = SumaTotalAncho
Next I
End Sub

Public Sub LeerArchivoPlano(DtaAux As Adodc, _
                            NombreFile As String)
Dim NumFile As Integer
Dim NumPos As Long
Dim RutaGeneraFile As String
Dim LineaTexto As String
RatonReloj
RutaGeneraFile = LeftStrg(CurDir$, 2) & "\SYSBASES\" & NombreFile
'MsgBox NombreFile
NumFile = FreeFile
Open RutaGeneraFile For Input As #NumFile ' Abre el archivo.
Line Input #NumFile, sSQL
Select_Adodc DtaAux, sSQL, False
FAConLineas = True
With DtaAux.Recordset
     ReDim TipoC(.fields.Count - 1) As Campos_Tabla
     Line Input #NumFile, LineaTexto
     I = 0
     Do
       TipoC(I).Campo = ObtenerCampoTexto(LineaTexto)
       LineaTexto = MidStrg(LineaTexto, Len(TipoC(I).Campo) + 2, Len(LineaTexto))
       I = I + 1
     Loop Until Len(LineaTexto) < 1
     Line Input #NumFile, LineaTexto
     I = 0
     Do
       TipoC(I).Ancho = Val(ObtenerCampoTexto(LineaTexto))
       LineaTexto = MidStrg(LineaTexto, Len(CStr(TipoC(I).Ancho)) + 2, Len(LineaTexto))
       I = I + 1
     Loop Until Len(LineaTexto) < 1
      
     Do While Not EOF(NumFile)
        Line Input #NumFile, LineaTexto
        NumPos = 1
        For I = 0 To .fields.Count - 1
            TipoC(I).Valor = SinEspaciosCampo(MidStrg(LineaTexto, NumPos, TipoC(I).Ancho))
            NumPos = NumPos + TipoC(I).Ancho
        Next I
       .AddNew
        For I = 0 To .fields.Count - 1
           .fields(TipoC(I).Campo) = TipoC(I).Valor
        Next I
       .Update
     Loop
End With
Close
RatonNormal
End Sub

Public Sub PrinterFields(Xo As Single, _
                         Yo As Single, _
                         AdoTipo As ADODB.Field, _
                         Optional PonerLineas As Boolean, _
                         Optional ImpLineaCero As Boolean)
'If Yo <= LimiteAlto Then
If ((Xo > 0) And (Yo > 0)) Then
   Distancia = CampoWidth(AdoTipo)
   If StrgFormatoCampo = Ninguno Then
      StrgFormatoCampo = " "
   ElseIf StrgFormatoCampo = "0" Or StrgFormatoCampo = "0.00" Then
      If Not ImpLineaCero Then StrgFormatoCampo = " "
   End If
   'PonerLineas
   If Yo < LimiteAlto Then
      
      LimpiarLinea Xo, Yo, PonerLineas
      Printer.CurrentX = Xo + Distancia
      Printer.CurrentY = Yo
      Printer.Print StrgFormatoCampo
   End If
  'MsgBox AdoTipo.Name & vbCrLf & StrgFormatoCampo
End If
'End If
End Sub

Public Sub EncabezadoDataReporte(Datas As Adodc, _
                                 X0 As Single, _
                                 x1 As Single)
Dim InicX As Single
Dim InicY As Single
Dim Y0 As Single
Dim y1 As Single
Dim LenT As Single
   PosLinea = 0
   PorteLetra = Printer.FontSize
   LetraAnterior = Printer.FontName
   Printer.FontName = TipoTimes
   Printer.FontSize = 9
   If X0 < 0.5 Then X0 = 0.5
   If x1 > LimiteAncho Then x1 = LimiteAncho
   If Printer.Orientation = 1 Then
      Y0 = 0: y1 = 1.6
   Else
      Y0 = 1: y1 = 2.6
   End If
   Printer.Line (X0, Y0)-(x1, y1), Negro, B
   Printer.Line (X0 + 0.2, Y0 + 0.2)-(X0 + 4.5, Y0 + 0.7), Negro, B
   Printer.Line (X0 + 0.2, Y0 + 0.9)-(X0 + 5, Y0 + 1.4), Negro, B
   Printer.Line (x1 - 4.5, Y0 + 0.2)-(x1 - 0.2, Y0 + 0.7), Negro, B
   Printer.Line (x1 - 4.5, Y0 + 0.9)-(x1 - 0.2, Y0 + 1.4), Negro, B
   Printer.Line (X0, y1)-(x1, y1 + 1.6), Negro, B
   Printer.FontBold = True
   Printer.FontSize = 9
   Printer.CurrentX = X0 + 0.4
   Printer.CurrentY = Y0 + 0.3
   Printer.Print "Fecha:"
   Printer.CurrentX = X0 + 0.4
   Printer.CurrentY = Y0 + 1
   Printer.Print "Usuario:"
   Printer.CurrentX = x1 - 4.2
   Printer.CurrentY = Y0 + 0.3
   Printer.Print "Hora:"
   Printer.CurrentX = x1 - 4.2
   Printer.CurrentY = Y0 + 1
   Printer.Print "Pagina No."
   Printer.FontSize = 20
   Printer.CurrentX = CentrarTextoEncab(Empresa, X0, x1): Printer.CurrentY = Y0 + 0.2
   Printer.Print Empresa
   Printer.FontBold = False
   Printer.FontSize = 9
   Printer.CurrentX = X0 + 1.6
   Printer.CurrentY = Y0 + 0.3
   Printer.Print date
   Printer.CurrentX = X0 + 1.7
   Printer.CurrentY = Y0 + 1
   Printer.Print NombreUsuario
   Printer.CurrentX = x1 - 2.8
   Printer.CurrentY = Y0 + 0.3
   Printer.Print Time
   Printer.CurrentX = x1 - 2.4
   Printer.CurrentY = Y0 + 1
   Printer.Print Pagina
   Printer.FontBold = True
   PosLinea = Y0 + 1.5
   With Heads
    Printer.FontSize = 12
    If .MsgTitulo <> "" Then
        PrinterTexto CentrarTexto(.MsgTitulo), PosLinea + 0.2, .MsgTitulo
        PosLinea = PosLinea + 0.8
    End If
    Printer.FontSize = 10
    If .MsgObjetivo <> "" And .TextoObjetivo <> "" Then
        PrinterTexto X0 + 0.3, PosLinea, .MsgObjetivo
        LenT = Printer.TextWidth(.MsgObjetivo)
        PrinterTexto X0 + LenT + 0.4, PosLinea, .TextoObjetivo
        PosLinea = PosLinea + 0.5
    End If
    If .MsgConcepto <> "" And .TextoConcepto <> "" Then
        PrinterTexto X0 + 0.3, PosLinea, .MsgConcepto
        LenT = Printer.TextWidth(.MsgConcepto)
        PrinterTexto X0 + LenT + 0.4, PosLinea, .TextoConcepto
        PosLinea = PosLinea + 0.6
    End If
   End With
PosLinea = PosLinea + 0.1
Printer.FontSize = 9
'========================================================================
PrinterAllFields CantCampos, PosLinea, Datas, True, True
PosLinea = PosLinea + 0.5
Printer.FontBold = False
Pagina = Pagina + 1
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
End Sub

Public Sub EncabezadoMeses(Datas As Adodc)
Dim InicX As Single
Dim InicY As Single
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
InicX = 0.5: InicY = 0.1
PrinterPaint LogoTipo, 0.5, 0, 3, 2
Printer.FontSize = 10
Cadena = "Página No. " & Pagina & "."
PrinterTexto TextoDerecha(Cadena) - 1, InicY, Cadena
Cadena = "Fecha: " & date$ & "."
PrinterTexto TextoDerecha(Cadena) - 1, InicY + 1, Cadena
PosLinea = InicY: Printer.FontBold = True: Printer.FontSize = 24
PrinterTexto CentrarTexto(Empresa), PosLinea, Empresa
PosLinea = PosLinea + 0.9: Printer.FontSize = 12
PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
PosLinea = PosLinea + 0.5
PrinterTexto CentrarTexto(SQLMsg2), PosLinea, SQLMsg2
Printer.FontSize = 9: PosLinea = 2.2
'========================================================================
For I = 0 To 14
   PrinterTexto Ancho(I), PosLinea, Datas.Recordset.fields(I).Name
Next I
LimpiarLinea Ancho(14), PosLinea, True
Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(15), PosLinea - 0.1), Negro
Printer.Line (Ancho(0), PosLinea + 0.4)-(Ancho(15), PosLinea + 0.4), Negro
PosLinea = PosLinea + 0.5
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
Printer.FontBold = False
End Sub

Public Sub EncabezadoMesesPorc(Datas As Adodc, _
                               FechaI As MaskEdBox, _
                               FechaF As MaskEdBox)
Dim InicX As Single
Dim InicY As Single
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
InicX = 0.5: InicY = 0.1
PrinterPaint LogoTipo, 0.5, 0, 3, 2
Printer.FontSize = 10
Cadena = "Página No. " & Pagina & "."
PrinterTexto TextoDerecha(Cadena) - 1, InicY, Cadena
Cadena = "Fecha: " & date$ & "."
PrinterTexto TextoDerecha(Cadena) - 1, InicY + 1, Cadena
PosLinea = InicY: Printer.FontBold = True: Printer.FontSize = 24
PrinterTexto CentrarTexto(Empresa), PosLinea, Empresa
PosLinea = PosLinea + 0.9: Printer.FontSize = 12
PrinterTexto CentrarTexto(SQLMsg1), PosLinea, SQLMsg1
PosLinea = PosLinea + 0.5
PrinterTexto CentrarTexto(SQLMsg2), PosLinea, SQLMsg2
Printer.FontSize = 9: PosLinea = 2.2
'========================================================================
With Datas.Recordset
     Printer.FontSize = 12
     Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(0), PosLinea + 0.8), Negro
     PrinterTexto Ancho(0), PosLinea, "C U E N T A S"
     Printer.FontSize = 7.5
     IR = 2
     For NoMeses = Month(FechaI.Text) To Month(FechaF.Text)
         Printer.Line (Ancho(IR), PosLinea - 0.1)-(Ancho(IR), PosLinea + 0.8), Negro
         PrinterTexto Ancho(IR), PosLinea - 0.05, UCaseStrg(MesesLetras(NoMeses))
         PrinterTexto Ancho(IR), PosLinea + 0.4, "V a l o r"
         IR = IR + 1
         Printer.Line (Ancho(IR), PosLinea + 0.4)-(Ancho(IR), PosLinea + 0.8), Negro
         PrinterTexto Ancho(IR), PosLinea + 0.4, "  %"
         IR = IR + 1
     Next NoMeses
End With
Printer.Line (Ancho(IR), PosLinea - 0.1)-(Ancho(IR), PosLinea + 0.8), Negro
PrinterTexto Ancho(IR), PosLinea - 0.05, "T O T A L"
PrinterTexto Ancho(IR), PosLinea + 0.4, "V a l o r"
IR = IR + 1
Printer.Line (Ancho(IR), PosLinea + 0.4)-(Ancho(IR), PosLinea + 0.8), Negro
PrinterTexto Ancho(IR), PosLinea + 0.4, "  %"
Printer.Line (Ancho(CantCampos), PosLinea - 0.1)-(Ancho(CantCampos), PosLinea + 0.9), Negro
Printer.Line (Ancho(0), PosLinea - 0.1)-(Ancho(CantCampos), PosLinea - 0.1), Negro
Printer.Line (Ancho(2), PosLinea + 0.35)-(Ancho(CantCampos), PosLinea + 0.35), Negro
Printer.Line (Ancho(0), PosLinea + 0.8)-(Ancho(CantCampos), PosLinea + 0.8), Negro
PosLinea = PosLinea + 0.9
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
Printer.FontBold = False
End Sub

Public Sub EncabezadoData(Datas As Adodc, _
                          Optional SegundaPagina As Boolean, _
                          Optional TCantCampos As Integer)
Dim InicX As Single
Dim InicY As Single
Dim TempCantCampos As Integer
TempCantCampos = CantCampos
If TCantCampos > 0 Then CantCampos = TCantCampos
Encabezado Ancho(0), AnchoPapel
PorteLetra = Printer.FontSize
LetraAnterior = Printer.FontName
Printer.FontName = TipoTimes
Printer.FontBold = False
If SQLMsg1 <> "" Then
   Printer.FontSize = 12
   PrinterTexto Ancho(0), PosLinea, SQLMsg1
   PosLinea = PosLinea + 0.45
End If
If SQLMsg2 <> "" Then
   Printer.FontSize = 10
   PrinterTexto Ancho(0), PosLinea, SQLMsg2
   PosLinea = PosLinea + 0.4
End If
If SQLMsg3 <> "" Then
   Printer.FontSize = 8
   PrinterTexto Ancho(0), PosLinea, SQLMsg3
   PosLinea = PosLinea + 0.4
End If
PosLinea = PosLinea + 0.05
Printer.FontSize = 8
Printer.FontBold = True
PrinterAllFields CantCampos, PosLinea, Datas, False, True, SegundaPagina
Printer.FontBold = False
PosLinea = PosLinea + 0.4
Printer.FontSize = PorteLetra
Printer.FontName = LetraAnterior
CantCampos = TempCantCampos
End Sub


Public Sub SelectDB_Combo(DBCombos As DataCombo, _
                          DataSQL As Adodc, _
                          SQLs As String, _
                          NombreCampo As String, _
                          Optional Final As Boolean, _
                          Optional NombreFile As String)
  If SQLs <> "" Then
     SQLs = CompilarSQL(SQLs)
    'MsgBox SQLs
     Generar_File_SQL NombreFile, SQLs
     DataSQL.RecordSource = SQLs
     DataSQL.Refresh
     DBCombos.ListField = DataSQL.Recordset.fields(NombreCampo).Name
     If DataSQL.Recordset.RecordCount > 0 Then
        If Final Then DataSQL.Recordset.MoveLast
        DBCombos.Text = DataSQL.Recordset.fields(NombreCampo)
        MarcarTexto DBCombos
     Else
        DBCombos.Text = "No existen datos."
     End If
  End If
End Sub

Public Sub SelectDB_List(DBLists As DataList, _
                         DataSQL As Adodc, _
                         SQLs As String, _
                         NombreCampo As String, _
                         Optional Final As Boolean, _
                         Optional NombreFile As String)
  If SQLs <> "" Then
     SQLs = CompilarSQL(SQLs)
    'MsgBox SQLs
     Generar_File_SQL NombreFile, SQLs
     DataSQL.RecordSource = SQLs
     DataSQL.Refresh
     DBLists.ListField = DataSQL.Recordset.fields(NombreCampo).Name
     If DataSQL.Recordset.RecordCount > 0 Then
        If Final Then DataSQL.Recordset.MoveLast
        DBLists.Text = DataSQL.Recordset.fields(NombreCampo)
     Else
        DBLists.Text = "No existen datos."
     End If
  End If
End Sub

Public Function DeleteSiNo(DataSQL As Adodc) As Boolean
Dim Cancelar As Boolean
  With DataSQL.Recordset
     Mensajes = "Eliminar la transacción: " & vbCrLf
     For I = 0 To .fields.Count - 1
        Mensajes = Mensajes & Space(20) & UCaseStrg(.fields(I).Name) & ": " & .fields(I) & "." & vbCrLf
     Next I
  End With
  Mensajes = Mensajes & vbCrLf _
           & "Realmente desea eliminar esta Transacción." & Space(10)
  Titulo = "Confirmación de Eliminación"
  If BoxMensaje = vbYes Then Cancelar = False Else Cancelar = True
  DeleteSiNo = Cancelar
End Function


Public Function ActualizarSiNo(DataSQL As Adodc) As Integer
Dim ResultadoSiNo As Integer
  Mensajes = "Actualizar la Transaccion: " & vbCrLf
  With DataSQL.Recordset
     For I = 0 To .fields.Count - 1
        Mensajes = Mensajes & Space(20) _
                 & UCaseStrg(.fields(I).Name) & ": " _
                 & .fields(I) & "." & vbCrLf
     Next I
  End With
  Mensajes = Mensajes & vbCrLf _
           & "Realmente desea Actualizar esta Transacción."
  Titulo = "Confirmación de Actualización"
  ActualizarSiNo = BoxMensaje
End Function

Public Sub CopyDataTemp(DatasFile As String)
Dim CopyOrigen As String
Dim CopyDestino As String
  CopyOrigen = UCaseStrg(RutaEmpresa & "\" & DatasFile)
  CopyDestino = UCaseStrg(RutaSubDirTemp & "\" & DatasFile)
  'Kill CopyDestino
  FileCopy CopyOrigen, CopyDestino
End Sub

Public Sub GenerarArchivoPlano(MiFrom As Form, _
                               DtaAux As Adodc, _
                               NombreFile As String, _
                               Optional DetFile As Boolean)
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim CaptionOld As String
Dim ValorBool As String
RatonReloj
CaptionOld = MiFrom.Caption
RutaGeneraFile = LeftStrg(RutaSysBases, 2) & "\SYSBASES\TEMP\" & NombreFile
'MsgBox NombreFile
NumFile = FreeFile
Contador = 0
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
FAConLineas = True
With DtaAux.Recordset
  If DetFile Then
     'Print #NumFile, DtaAux.RecordSource
     ReDim TipoC(.fields.Count - 1) As Campos_Tabla
     For I = 0 To .fields.Count - 1
         TipoC(I).Campo = .fields(I).Name
         TipoC(I).Ancho = AnchoTipoCampoTexto(.fields(I))
     Next I
     For I = 0 To .fields.Count - 1
         Cadena2 = TipoC(I).Campo
         'MsgBox UCaseStrg(Cadena2)
         If UCaseStrg(Cadena2) = "ID" Then Cadena2 = "I"
         If I <> .fields.Count - 1 Then
            Print #NumFile, SetearBlancos(Cadena2, Len(TipoC(I).Campo), 0, False, FAConLineas);
         Else
            Print #NumFile, SetearBlancos(Cadena2, Len(TipoC(I).Campo), 0, False, FAConLineas)
         End If
     Next I
     For I = 0 To .fields.Count - 1
      Cadena = CStr(TipoC(I).Ancho)
      If I <> .fields.Count - 1 Then
         Print #NumFile, SetearBlancos(Cadena, Len(Cadena), 0, True, FAConLineas);
      Else
         Print #NumFile, SetearBlancos(Cadena, Len(Cadena), 0, True, FAConLineas)
      End If
     Next I
  Else
     ReDim TipoC(.fields.Count - 1) As Campos_Tabla
     For I = 0 To .fields.Count - 1
         TipoC(I).Campo = .fields(I).Name
         TipoC(I).Ancho = AnchoTipoCampoTexto(.fields(I))
     Next I
  End If
     FAConLineas = True
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
         MiFrom.Caption = RutaGeneraFile & ": Registro No. " & Format$(Contador / .RecordCount, "00%")
         For I = 0 To .fields.Count - 1
          If I <> .fields.Count - 1 Then
             Select Case .fields(I).Type
               Case TadByte, TadInteger, TadLong
                    Print #NumFile, SetearBlancos(.fields(I), TipoC(I).Ancho, 0, True, FAConLineas);
               Case TadSingle, TadDouble, TadCurrency
                    Print #NumFile, SetearBlancos(.fields(I), TipoC(I).Ancho, 0, True, FAConLineas, True);
               Case TadBoolean
                    If .fields(I) = 0 Then
                        ValorBool = "0"
                    Else
                        ValorBool = "1"
                    End If
                    Print #NumFile, SetearBlancos(ValorBool, TipoC(I).Ancho, 0, True, FAConLineas);
               Case Else
                    Print #NumFile, SetearBlancos(.fields(I), TipoC(I).Ancho, 0, False, FAConLineas);
             End Select
          Else
             Select Case .fields(I).Type
               Case TadByte, TadInteger, TadLong
                    Print #NumFile, SetearBlancos(.fields(I), TipoC(I).Ancho, 0, True, FAConLineas)
               Case TadSingle, TadDouble, TadCurrency
                    Print #NumFile, SetearBlancos(.fields(I), TipoC(I).Ancho, 0, True, FAConLineas, True)
               Case TadBoolean
                    If .fields(I) = 0 Then
                        ValorBool = "0"
                    Else
                        ValorBool = "1"
                    End If
                    Print #NumFile, SetearBlancos(ValorBool, TipoC(I).Ancho, 0, True, FAConLineas)
               Case Else
                    Print #NumFile, SetearBlancos(.fields(I), TipoC(I).Ancho, 0, False, FAConLineas)
             End Select
          End If
         Next I
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

Public Sub GenerarTablaArchivoPlano(MiForm As Form, _
                                    FechaResp As String, _
                                    FechaDesde As String, _
                                    FechaHasta As String, _
                                    NombreTabla As String, _
                                    DtaAux As Adodc)
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim CaptionOld As String
Dim NombreFile As String
Dim CadFileReg As String
Dim ContadorReg As Long
Dim TotalCampo As Integer
Dim ValorBool As String
RatonReloj
ContadorReg = 0
If FileResp <= 0 Then FileResp = 1
'MsgBox NombreFile & " .............."
With DtaAux.Recordset
 If .RecordCount > 0 Then
    .MoveLast
     TotalReg = .RecordCount
     TotalCampo = .fields.Count - 1
     Progreso_Barra.Incremento = 0
     Progreso_Barra.Valor_Maximo = TotalReg
     Progreso_Barra.Mensaje_Box = NombreTabla
     Progreso_Esperar
    .MoveFirst
     NombreFile = "F" & Format$(Day(FechaResp), "00") _
                & Format$(Month(FechaResp), "00") _
                & Format$(FileResp, "000") & ".TXT"
     TextoFileEmp = TextoFileEmp & vbCrLf & NombreFile & " => " & NombreTabla
     RutaGeneraFile = LeftStrg(RutaSysBases, 2) & "\SYSBASES\DATOS\D" & GrupoEmpresa & "\" & NombreFile
     MiForm.LstArchivo.AddItem NombreFile & " => " & NombreTabla
     MiForm.LstArchivo.Refresh
     NumFile = FreeFile
     Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
     FAConLineas = True
     Print #NumFile, Format$(TotalReg, "##0") & " - " & FechaDesde & " - " & FechaHasta & " - " & NombreTabla
     ReDim TipoC(TotalCampo) As Campos_Tabla
     For I = 0 To TotalCampo
         TipoC(I).Campo = CompilarString(.fields(I).Name)
         TipoC(I).Ancho = AnchoTipoCampoTexto(.fields(I))
     Next I
     CadFileReg = ""
     For I = 0 To TotalCampo
         CadFileReg = CadFileReg & TipoC(I).Campo & "|"
     Next I
     Print #NumFile, CadFileReg
     FAConLineas = True
    .MoveFirst
     Do While Not .EOF
        ContadorReg = ContadorReg + 1
        Progreso_Esperar
        CadFileReg = ""
        For I = 0 To TotalCampo
          If IsNull(.fields(I)) Or IsEmpty(.fields(I)) Then Codigo4 = "0" Else Codigo4 = CStr(.fields(I))
          Select Case .fields(I).Type
            Case TadBoolean
                 If Codigo4 = Ninguno Then Codigo4 = "0"
                 Codigo4 = CStr(CInt(CBool(Codigo4)))
                 Codigo4 = SetearBlancos(Codigo4, 2, 0, True, FAConLineas)
            Case TadByte, TadInteger, TadLong
                 If .fields(I).Name = "Item" Then
                     Codigo4 = SetearBlancos(Format$(.fields(I), "000"), 3, 0, False, FAConLineas)
                 Else
                     Codigo4 = SetearBlancos(Codigo4, 0, 0, True, FAConLineas)
                 End If
            Case TadSingle, TadDouble, TadCurrency
                 Codigo4 = SetearBlancos(Codigo4, 0, 0, True, FAConLineas, True)
            Case TadText
                 If UCaseStrg(.fields(I).Name) = "RUC_CI" Then Codigo4 = CompilarRUC_CI(.fields(I))
                 If UCaseStrg(.fields(I).Name) = "RUC" Then Codigo4 = CompilarRUC_CI(.fields(I))
                 If UCaseStrg(.fields(I).Name) = "CI" Then Codigo4 = CompilarRUC_CI(.fields(I))
                 If UCaseStrg(.fields(I).Name) = "CEDULA" Then Codigo4 = CompilarRUC_CI(.fields(I))
                 If Codigo4 = "0" Then Codigo4 = Ninguno
                 If Codigo4 = " " Then Codigo4 = Ninguno
                 Codigo4 = SetearBlancos(Codigo4, 0, 0, False, FAConLineas)
            Case Else
                 If Codigo4 = "0" Then Codigo4 = Ninguno
                 If Codigo4 = " " Then Codigo4 = Ninguno
                 Codigo4 = SetearBlancos(Codigo4, 0, 0, False, FAConLineas)
           End Select
           Codigo4 = Replace(Codigo4, vbCrLf, Chr(17) & Chr(31))
           Codigo4 = Replace(Codigo4, "'", "`")
           CadFileReg = CadFileReg & Codigo4
        Next I
        Print #NumFile, CadFileReg
       .MoveNext
     Loop
    .MoveFirst
     FileResp = FileResp + 1
 End If
End With
Close #NumFile
Progreso_Final
RatonNormal
End Sub

'''Public Sub GenerarArchivoCOBOL(MiForm As Form, _
'''                               FechaResp As String, _
'''                               NombreTabla As String, _
'''                               DtaAux As Adodc)
'''Dim NumFile As Integer
'''Dim RutaGeneraFile As String
'''Dim CaptionOld As String
'''Dim NombreFile As String
'''Dim ContadorReg As Long
'''Dim ValorBool As String
'''RatonReloj
'''ContadorReg = 0
'''If FileResp <= 0 Then FileResp = 1
'''With DtaAux.Recordset
''' If .RecordCount > 0 Then
'''    .MoveLast
'''     TotalReg = .RecordCount
'''    .MoveFirst
'''     NombreFile = NombreTabla & ".TXT"
'''     RutaGeneraFile = LeftStrg(RutaSysBases, 2) & "\SYSBASES\" & NombreFile
'''     MiForm.TextArchivo.Text = MiForm.TextArchivo.Text & Space(7) & NombreFile & " => " & NombreTabla & vbCrLf
'''     MiForm.TextArchivo.Refresh
'''    'MsgBox NombreFile
'''     NumFile = FreeFile
'''     'Contador = 0
'''     Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
'''     FAConLineas = True
'''     Print #NumFile, Format$(TotalReg, "##0") & " - " & NombreTabla
'''     ReDim TipoC(.Fields.Count - 1) As Campos_Tabla
'''     For I = 0 To .Fields.Count - 1
'''         TipoC(I).Campo = .Fields(I).Name
'''         TipoC(I).Ancho = AnchoTipoCampoTexto(.Fields(I))
'''     Next I
'''     For I = 0 To .Fields.Count - 1
'''         Cadena2 = CompilarString(TipoC(I).Campo) & "|"
'''         If I <> .Fields.Count - 1 Then
'''            Print #NumFile, Cadena2;
'''         Else
'''            Print #NumFile, Cadena2
'''         End If
'''     Next I
'''     ReDim TipoC(.Fields.Count - 1) As Campos_Tabla
'''     For I = 0 To .Fields.Count - 1
'''         TipoC(I).Campo = .Fields(I).Name
'''         TipoC(I).Ancho = AnchoTipoCampoTexto(.Fields(I))
'''     Next I
'''     FAConLineas = True
'''    .MoveFirst
'''     Do While Not .EOF
'''        ContadorReg = ContadorReg + 1
'''        MiForm.Caption = NombreTabla & ": Procesando(" & Format$(ContadorReg / TotalReg, "##0.00%") & ") " & String$(ContadorReg Mod 40, "|")
'''        For I = 0 To .Fields.Count - 1
'''          If IsNull(.Fields(I)) Or IsEmpty(.Fields(I)) Then
'''             Codigo4 = "0"
'''          Else
'''             Codigo4 = CStr(.Fields(I))
'''          End If
'''          If I <> .Fields.Count - 1 Then
'''             Select Case .Fields(I).Type
'''               Case TadByte, TadInteger, TadLong
'''                    If .Fields(I).Name = "Item" Then
'''                        Codigo4 = Format$(.Fields(I), "000")
'''                        Print #NumFile, SetearBlancos(Codigo4, 3, 0, False, FAConLineas);
'''                    Else
'''                        Print #NumFile, SetearBlancos(Codigo4, 14, 0, True, FAConLineas);
'''                    End If
'''               Case TadSingle, TadCurrency, TadDouble
'''                    Print #NumFile, SetearBlancos(Codigo4, 14, 0, True, FAConLineas, True);
'''               Case TadDate, TadDate1
'''                    Codigo4 = Format$(.Fields(I), "YYYYMMDD")
'''                    Print #NumFile, SetearBlancos(Codigo4, 8, 0, True, FAConLineas);
'''               Case TadBoolean
'''                    Codigo4 = CStr(CInt(CBool(Codigo4)))
'''                    Print #NumFile, SetearBlancos(Codigo4, 2, 0, True, FAConLineas);
'''               Case Else
'''                    If UCaseStrg(.Fields(I).Name) = "RUC_CI" Then Codigo4 = CompilarRUC_CI(.Fields(I))
'''                    If UCaseStrg(.Fields(I).Name) = "RUC" Then Codigo4 = CompilarRUC_CI(.Fields(I))
'''                    If UCaseStrg(.Fields(I).Name) = "CI" Then Codigo4 = CompilarRUC_CI(.Fields(I))
'''                    If UCaseStrg(.Fields(I).Name) = "CEDULA" Then Codigo4 = CompilarRUC_CI(.Fields(I))
'''                    Print #NumFile, SetearBlancos(Codigo4, .Fields(I).DefinedSize, 0, False, FAConLineas);
'''             End Select
'''          Else
'''             Select Case .Fields(I).Type
'''               Case TadByte, TadInteger, TadLong
'''                    If .Fields(I).Name = "Item" Then
'''                        Codigo4 = Format$(.Fields(I), "000")
'''                        Print #NumFile, SetearBlancos(Codigo4, 3, 0, False, FAConLineas)
'''                    Else
'''                        Print #NumFile, SetearBlancos(Codigo4, 14, 0, True, FAConLineas)
'''                    End If
'''               Case TadSingle, TadCurrency, TadDouble
'''                    Print #NumFile, SetearBlancos(Codigo4, 14, 0, True, FAConLineas, True)
'''               Case TadDate, TadDate1
'''                    Codigo4 = Format$(.Fields(I), "YYYYMMDD")
'''                    Print #NumFile, SetearBlancos(Codigo4, 8, 0, True, FAConLineas)
'''               Case TadBoolean
'''                    Codigo4 = CStr(CInt(CBool(Codigo4)))
'''                    Print #NumFile, SetearBlancos(Codigo4, 2, 0, True, FAConLineas)
'''               Case Else
'''                    If UCaseStrg(.Fields(I).Name) = "RUC_CI" Then Codigo4 = CompilarRUC_CI(.Fields(I))
'''                    If UCaseStrg(.Fields(I).Name) = "RUC" Then Codigo4 = CompilarRUC_CI(.Fields(I))
'''                    If UCaseStrg(.Fields(I).Name) = "CI" Then Codigo4 = CompilarRUC_CI(.Fields(I))
'''                    If UCaseStrg(.Fields(I).Name) = "CEDULA" Then Codigo4 = CompilarRUC_CI(.Fields(I))
'''                    Print #NumFile, SetearBlancos(Codigo4, .Fields(I).DefinedSize, 0, False, FAConLineas)
'''             End Select
'''          End If
'''        Next I
'''       .MoveNext
'''     Loop
'''    .MoveFirst
'''     FileResp = FileResp + 1
''' End If
'''End With
'''Close #NumFile
'''RatonNormal
'''End Sub

Public Sub ImprimirDataObjetivo(Datas As Adodc, _
                                FinDoc As Boolean, _
                                FormaImp As Byte, _
                                SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
'Escala_Centimetro FormaImp, TipoTimes, SizeLetra
DataAnchoCampos InicioX, Datas, SizeLetra, TipoTimes, FormaImp
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
     .MoveFirst
      EncabezadoDataReporte Datas, Ancho(0), Ancho(CantCampos)
      Printer.FontSize = SizeLetra
      Do While Not .EOF
         PrinterAllFields CantCampos, PosLinea, Datas, True, False
         PosLinea = PosLinea + 0.4
         If PosLinea > LimiteAlto Then
            Printer.Line (Ancho(0), PosLinea)-(Ancho(CantCampos), PosLinea), Negro
            Printer.NewPage
            EncabezadoDataReporte Datas, Ancho(0), Ancho(CantCampos)
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
      Loop
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirMeses(Datas As Adodc, _
                         SizeLetra As Integer)
On Error GoTo Errorhandler
RatonReloj
InicioX = 0.5: InicioY = 0
Escala_Centimetro 2, TipoTimes, SizeLetra
Pagina = 1
CantCampos = 15
ReDim Ancho(CantCampos + 1) As Single
Ancho(0) = 0.5   'Codogo
Ancho(1) = 0.5   'Detalle
Ancho(2) = 3.9   'Ene
Ancho(3) = 5.6   'Feb
Ancho(4) = 7.3   'Mar
Ancho(5) = 9     'Abr
Ancho(6) = 10.7  'May
Ancho(7) = 12.4  'Jun
Ancho(8) = 14.1  'Jul
Ancho(9) = 15.8  'Ago
Ancho(10) = 17.5 'Sep
Ancho(11) = 19.2 'Oct
Ancho(12) = 20.9 'Nov
Ancho(13) = 22.6 'Dic
Ancho(14) = 24.3 'TOTAL
Ancho(15) = 26
'Iniciamos la impresion
PosLinea = 0
With Datas.Recordset
    .MoveFirst
     EncabezadoMeses Datas
     Printer.FontBold = False
     Printer.FontSize = SizeLetra
     'Codigo = Datas.Recordset.Fields(0)
     Do While Not .EOF
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.35
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
           Printer.NewPage
           PosLinea = 0
           EncabezadoMeses Datas
           Printer.FontSize = SizeLetra
        End If
       .MoveNext
     Loop
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
UltimaLinea = PosLinea + 0.5
MensajeEncabData = ""
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub Imprimir12MesesPorc(Datas As Adodc, _
                               FechaI As MaskEdBox, _
                               FechaF As MaskEdBox)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
RatonReloj
FechaValida FechaI
FechaValida FechaF
SizeLetra = 5
InicioX = 0.1: InicioY = 0
Pagina = 1
CantCampos = Datas.Recordset.fields.Count
ReDim Ancho(CantCampos + 2) As Single
Ancho(0) = 0.1   'Codogo
Ancho(1) = 0.1   'Detalle
Ancho(2) = 3     'Inicio de meses con porcentaje
Distancia = 3
NoMeses = Month(FechaF.Text) - Month(FechaI.Text) + 1
'If NoMeses > 6 Then NoMeses = 6
IE = 2: KE = NoMeses
For JE = 1 To NoMeses
    Ancho(IE) = Distancia       ' Mes
    Distancia = Redondear(Distancia + 1.2, 2)
    IE = IE + 1
    Ancho(IE) = Distancia       ' Mes%
    Distancia = Redondear(Distancia + 0.7, 2)
    IE = IE + 1
Next JE
Ancho(IE) = Distancia 'TOTAL
Distancia = Redondear(Distancia + 1.2, 2)
IE = IE + 1
Ancho(IE) = Distancia 'TOT %
Distancia = Redondear(Distancia + 0.7, 2)
IE = IE + 1
Ancho(IE) = Distancia 'Fin
'Iniciamos la impresion
If Distancia <= 19.5 Then
   Escala_Centimetro 1, TipoTimes, SizeLetra
Else
   Escala_Centimetro 2, TipoTimes, SizeLetra, True
End If
IE = 1: JE = IE
Volver_Imp:
PosLinea = 0
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
    'Encabezado Ancho(0), Ancho(IR)
     EncabezadoMesesPorc Datas, FechaI, FechaF
     Printer.FontBold = False
     Printer.FontSize = SizeLetra
    'Codigo = Datas.Recordset.Fields(0)
     Do While Not .EOF
        Printer.FontSize = SizeLetra
        Printer.FontBold = True
        'PrinterFields Ancho(1), PosLinea, .Fields("Cuenta")
        'NoDias = 2
        ' For KE = IE To JE
        '    PrinterFields Ancho(NoDias), PosLinea, .Fields(MesesLetras(KE))
        '    NoDias = NoDias + 1
        'Next KE
        'If NoMeses <= 12 Then
        '   PrinterFields Ancho(NoDias), PosLinea, .Fields("TOTAL")
        '   NoDias = NoDias + 1
        '   PrinterFields Ancho(NoDias), PosLinea, .Fields("TOT_P")
        'End If
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
         PosLinea = PosLinea + 0.3
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            Printer.NewPage
            PosLinea = 0
            'Encabezado Ancho(0), Ancho(IR)
            EncabezadoMesesPorc Datas, FechaI, FechaF
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
     Loop
  End If
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
UltimaLinea = PosLinea + 0.5
MensajeEncabData = ""
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub ImprimirMesesPorc(Datas As Adodc, _
                             FechaI As MaskEdBox, _
                             FechaF As MaskEdBox)
Dim SizeLetra As Integer
On Error GoTo Errorhandler
RatonReloj
FechaValida FechaI
FechaValida FechaF
SizeLetra = 7
InicioX = 0.1: InicioY = 0
Pagina = 1
CantCampos = Datas.Recordset.fields.Count
ReDim Ancho(CantCampos + 1) As Single
Ancho(0) = 0.1   'Codogo
Ancho(1) = 0.1   'Detalle
Ancho(2) = 4   'Inicio de meses con porcentaje
Distancia = 4
IR = 2
For NoMeses = Month(FechaI.Text) To Month(FechaF.Text)
    Distancia = Distancia + 1.7
    IR = IR + 1
    Ancho(IR) = Distancia
    Distancia = Distancia + 1.2
    IR = IR + 1
    Ancho(IR) = Distancia
Next NoMeses
Distancia = Distancia + 1.7
IR = IR + 1
Ancho(IR) = Distancia 'TOTAL
Distancia = Distancia + 1.2
IR = IR + 1
Ancho(IR) = Distancia 'TOT
'Iniciamos la impresion
If Distancia <= 19.5 Then
   Escala_Centimetro 1, TipoTimes, SizeLetra
Else
   Escala_Centimetro 2, TipoTimes, SizeLetra
End If
PosLinea = 0
With Datas.Recordset
    .MoveFirst
     'Encabezado Ancho(0), Ancho(IR)
     EncabezadoMesesPorc Datas, FechaI, FechaF
     Printer.FontBold = False
     Printer.FontSize = SizeLetra
     'Codigo = Datas.Recordset.Fields(0)
     Do While Not .EOF
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        ' LimpiarLinea Ancho(12), PosLinea, True
         PosLinea = PosLinea + 0.35
         If PosLinea >= LimiteAlto Then
            Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos)
            Printer.NewPage
            PosLinea = 0
            'Encabezado Ancho(0), Ancho(IR)
            EncabezadoMesesPorc Datas, FechaI, FechaF
            Printer.FontSize = SizeLetra
         End If
        .MoveNext
     Loop
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
UltimaLinea = PosLinea + 0.5
MensajeEncabData = ""
RatonNormal
Printer.EndDoc
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Sub GenerarDataTexto(MiForm As Form, _
                            Dtas As Adodc, _
                            Optional ConEncabezado As Boolean)
Dim EncabezadoHojaExcel As String
   If Dtas.Recordset.RecordCount > 0 Then
      EncabezadoHojaExcel = MiForm.Caption
      If DescripcionEstado = "" Then DescripcionEstado = "NO ESTA PERMITIDO"
      Exportar_AdoDB_Excel Dtas.Recordset, EncabezadoHojaExcel
   End If
End Sub

Public Sub Select_Adodc(DataSQL As Adodc, _
                        SQLs As String, _
                        Optional SetMsg As Boolean, _
                        Optional Decimales As String, _
                        Optional NombreFile As String)
Dim CantDecCampo As String
Dim Strgs As String
Dim LenCamposDec As Long
Dim CantDec As Long
Dim Idx As Long
Dim Kdx As Long
Dim Jdx As Long

   If SQLs <> "" Then
      RatonReloj
      SQLs = CompilarSQL(SQLs)
      Generar_File_SQL NombreFile, SQLs
     'MsgBox AdoStrCnn & vbCrLf & String(92, "-") & vbCrLf & SQLs
      DataSQL.RecordSource = SQLs
      DataSQL.Refresh
      
     'Determinamos el ancho de los campos
      CantCampos = DataSQL.Recordset.fields.Count
     
     'Array para almacenar el ancho de cada columna
      ReDim Vect_Dec(CantCampos) As Campos_Decimal
     'Enceramos la impresion
      For Idx = 0 To CantCampos - 1
          Vect_Dec(Idx).Campo = DataSQL.Recordset.fields(Idx).Name
          Vect_Dec(Idx).CantDec = 2
          Vect_Dec(Idx).AnchoCampo = 0
      Next Idx

      For Idx = 0 To CantCampos - 1
         'MsgBox Decimales & vbCrLf & Vect_Dec(Col).Campo & vbCrLf & Vect_Dec(Col).CantDec & vbCrLf & Vect_Dec(Col).AnchoCampo
          If Decimales <> "" Then
             LenCamposDec = Len(DataSQL.Recordset.fields(Idx).Name)
             For Jdx = 1 To Len(Decimales)
                 If DataSQL.Recordset.fields(Idx).Name = MidStrg(Decimales, Jdx, LenCamposDec) Then
                    CantDecCampo = ""
                    For CantDec = Jdx + LenCamposDec To Len(Decimales)
                        If MidStrg(Decimales, CantDec, 1) = "|" Then CantDec = Len(Decimales) + 1
                        CantDecCampo = CantDecCampo & MidStrg(Decimales, CantDec, 1)
                    Next CantDec
                    Vect_Dec(Idx).CantDec = Val(CantDecCampo)
                 End If
             Next Jdx
          End If
      Next Idx
     'If SQL_Server = False Then DataSQL.Refresh
      RatonNormal
      With DataSQL.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
           Strgs = "CONSULTA PROCESADA CORRECTAMENTE." & vbCrLf & vbCrLf _
                 & "Existen " & .RecordCount & " Registro(s) Procesado(s)."
           If SetMsg Then MsgBox Strgs
           Strgs = "Registros: " & Format$(.RecordCount, "#,##0") & "." _
                 & Space(30) & "Página(s): " _
                 & Format$((.RecordCount / 45) + 1, "#,##0") & "."
           DataSQL.Caption = Strgs
       Else
           Strgs = "No existen Datos Disponibles." & vbCrLf _
                 & "'Si cree tener datos, consulte al su Técnico'."
           If SetMsg Then MsgBox Strgs
       End If
      End With
   End If
End Sub

Public Sub Select_Adodc_Grid(DBGMalla As DataGrid, _
                             Dtas As Adodc, _
                             SQuerys As String, _
                             Optional Decimales As String, _
                             Optional EsCampoCorto As Boolean, _
                             Optional PresentarEsperar As Boolean, _
                             Optional NombreFile As String)
Dim AnchoMax As Integer
Dim Presentar As Boolean
Dim LenCamposDec As Integer
Dim CantDecCampo As String
Dim Otros_Dec As Boolean
Dim CantDec As Long
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

On Error GoTo Errorhandler
 
   If SQuerys <> "" Then
      RatonReloj
     'Determinamos el ancho de los tipo de campo que tiene VB
      Determina_Ancho_Tipos EsCampoCorto
      WidthBoolean = CSng(DBGMalla.Parent.TextWidth(CadBoolean))
      WidthDate = CSng(DBGMalla.Parent.TextWidth(UCaseStrg(CadDate)))
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
      Generar_File_SQL NombreFile, SQuerys
     'MsgBox AdoStrCnn
     'MsgBox SQuerys
      Dtas.RecordSource = SQuerys
      Dtas.Refresh

     'Variables para la cantidad de registros y columnas
      CantCampos = Dtas.Recordset.fields.Count
      
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
      DBGMalla.HeadFont.bold = True
      For Col = 0 To CantCampos - 1
          Vect_Dec(Col).Campo = Dtas.Recordset.fields(Col).Name
          Vect_Dec(Col).CantDec = 2
          Vect_Dec(Col).AnchoCampo = DBGMalla.Parent.TextWidth(UCase(Dtas.Recordset.fields(Col).Name & " "))
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
                Vect_Dec(Col).CantDec = Val(MidStrg(CadAux, C1, C2 - C1 + 1))
                CadAux = TrimStrg(MidStrg(CadAux, C2 + 2, Len(CadAux)))
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
               Select Case Dtas.Recordset.fields(Col).Type
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
                      CantDec = 2
                      C1 = InStr(Decimales, Dtas.Recordset.fields(Col).Name)
                      If C1 > 0 Then
                         CadAux = MidStrg(Decimales, C1, Len(Decimales))
                         C1 = InStr(CadAux, " ")
                         C2 = InStr(CadAux, "|")
                         CantDec = Val(MidStrg(CadAux, C1, C2 - C1))
                      End If
                      'MsgBox Dtas.Recordset.Fields(Col).Name & vbCrLf & CantDec
                     .Columns(Col).Alignment = dbgRight
                     .Columns(Col).NumberFormat = "#,##0." & String$(CantDec, "0")
                      If Dtas.Recordset.fields(Col).Type = TadDouble Then
                        If Vect_Dec(Col).AnchoCampo < WidthDouble Then Vect_Dec(Col).AnchoCampo = WidthDouble
                      Else
                        If Vect_Dec(Col).AnchoCampo < WidthCurrency Then Vect_Dec(Col).AnchoCampo = WidthCurrency
                      End If
                 Case TadText
                     'Determinamos el ancho maximo mas abajo dependendo de la cantidad de caracteres por celda
                      AnchoMax = Dtas.Recordset.fields(Col).DefinedSize
                      If AnchoMax > 50 Then AnchoMax = 50
                      If EsCampoCorto Then
                         WidthText = DBGMalla.Parent.TextWidth(String$(AnchoMax, ".") & " ")
                      Else
                         WidthText = DBGMalla.Parent.TextWidth(String$(AnchoMax, "H") & " ")
                      End If
                      If Vect_Dec(Col).AnchoCampo < WidthText Then Vect_Dec(Col).AnchoCampo = WidthText
                 Case Else
                      If Dtas.Recordset.fields(Col).DefinedSize <= 50 Then
                         AnchoMax = Dtas.Recordset.fields(Col).DefinedSize
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

'''Public Sub SelectMSFGrid(DBGMalla As MSHFlexGrid, _
'''                         Datas As Adodc, _
'''                         SQuerys As String, _
'''                         Optional Decimales As String, _
'''                         Optional EsCampoCorto As Boolean)
'''Dim CadFlexG As String
'''Dim NomCamp As String
'''Dim CadCamp As String
'''Dim CantDecCampo As String
'''Dim anchoFields As Single
'''Dim Presentar As Boolean
'''Dim LenCamposDec As Long
'''  On Error GoTo Errorhandler
'''  DBGMalla.Visible = False
'''  RatonReloj
'''  SetAnchoCampos DBGMalla, EsCampoCorto
'''  CadFlexG = " * |"
'''  SQuerys = CompilarSQL(SQuerys)
'''' MsgBox SQuerys
'''  Datas.RecordSource = SQuerys
'''  Datas.Refresh
''' 'Enceramos la impresion
'''  ReDim Vect_Dec(Datas.Recordset.Fields.Count) As Campos_Decimal
'''  For i = 0 To Datas.Recordset.Fields.Count
'''      Vect_Dec(i).Campo = "Ninguno"
'''      Vect_Dec(i).CantDec = 0
'''  Next i
'''  For i = 0 To Datas.Recordset.Fields.Count - 1
'''      'MsgBox Decimales & vbCrLf & Vect_Dec(I).Campo & vbCrLf & Vect_Dec(I).CantDec
'''      If Decimales <> "" Then
'''         LenCamposDec = Len(Datas.Recordset.Fields(i).Name)
'''         For J = 1 To Len(Decimales)
'''             If Datas.Recordset.Fields(i).Name = MidStrg(Decimales, J, LenCamposDec) Then
'''                Vect_Dec(i).Campo = Datas.Recordset.Fields(i).Name
'''                CantDecCampo = ""
'''                For K = J + LenCamposDec To Len(Decimales)
'''                    If MidStrg(Decimales, K, 1) = "|" Then K = Len(Decimales) + 1
'''                    CantDecCampo = CantDecCampo & MidStrg(Decimales, K, 1)
'''                Next K
'''                Vect_Dec(i).CantDec = Val(CantDecCampo)
'''             End If
'''         Next J
'''      End If
'''  Next i
''' 'Determinamos el ancho de los Campos
'''  With Datas.Recordset
'''   For i = 0 To .Fields.Count - 1
'''       NomCamp = .Fields(i).Name
'''       'MsgBox NomCamp & vbCrLf & CadFlexG
'''       Select Case .Fields(i).Type
'''         Case TadBoolean
'''              CadCamp = CadBoolean
'''              If Len(NomCamp) > Len(CadCamp) Then
'''                 CadFlexG = CadFlexG & "<" & NomCamp & "|"
'''              Else
'''                 CadFlexG = CadFlexG & "<" & NomCamp & String$(Len(CadCamp) - Len(NomCamp), " ") & "|"
'''              End If
'''         Case TadDate, TadDate1
'''              CadCamp = CadDate
'''              If Len(NomCamp) > Len(CadCamp) Then
'''                 CadFlexG = CadFlexG & "<" & NomCamp & "|"
'''              Else
'''                 CadFlexG = CadFlexG & "<" & NomCamp & String$(Len(CadCamp) - Len(NomCamp), " ") & "|"
'''              End If
'''         Case TadByte
'''              CadCamp = CadByte
'''              If Len(NomCamp) > Len(CadCamp) Then
'''                 CadFlexG = CadFlexG & ">" & NomCamp & "|"
'''              Else
'''                 CadFlexG = CadFlexG & ">" & NomCamp & String$(4 - Len(NomCamp), " ") & "|"
'''              End If
'''         Case TadInteger
'''              CadCamp = CadInteger
'''              If Len(NomCamp) > Len(CadCamp) Then
'''                 CadFlexG = CadFlexG & ">" & NomCamp & "|"
'''              Else
'''                 CadFlexG = CadFlexG & ">" & NomCamp & String$(7 - Len(NomCamp), " ") & "|"
'''              End If
'''         Case TadLong
'''              CadCamp = CadLong
'''              If Len(NomCamp) > Len(CadCamp) Then
'''                 CadFlexG = CadFlexG & ">" & NomCamp & "|"
'''              Else
'''                 CadFlexG = CadFlexG & ">" & NomCamp & String$(9 - Len(NomCamp), " ") & "|"
'''              End If
'''         Case TadSingle
'''              CadCamp = CadSingle
'''              If Len(NomCamp) > Len(CadCamp) Then
'''                 CadFlexG = CadFlexG & ">" & NomCamp & "|"
'''              Else
'''                 CadFlexG = CadFlexG & ">" & NomCamp & String$(9 - Len(NomCamp), " ") & "|"
'''              End If
'''         Case TadDouble
'''              CadCamp = CadDouble
'''              If Len(NomCamp) > Len(CadCamp) Then
'''                 CadFlexG = CadFlexG & ">" & NomCamp & "|"
'''              Else
'''                 CadFlexG = CadFlexG & ">" & NomCamp & String$(12 - Len(NomCamp), " ") & "|"
'''              End If
'''         Case TadCurrency
'''              CadCamp = CadCurrency
'''              If Len(NomCamp) > Len(CadCamp) Then
'''                 CadFlexG = CadFlexG & ">" & NomCamp & "|"
'''              Else
'''                 CadFlexG = CadFlexG & ">" & NomCamp & String$(12 - Len(NomCamp), " ") & "|"
'''              End If
'''         Case TadText
'''              CadCamp = String$(.Fields(i).DefinedSize + 5, " ")
'''              If Len(CadCamp) > 50 Then CadCamp = MidStrg(CadCamp, 1, 40)
'''              If Len(NomCamp) > Len(CadCamp) Then
'''                 CadFlexG = CadFlexG & "<" & NomCamp & "|"
'''              Else
'''                 CadFlexG = CadFlexG & "<" & NomCamp & String$(Len(CadCamp) - Len(NomCamp), " ") & "|"
'''              End If
'''         Case Else
'''              CadCamp = String$(3, " ")
'''              CadFlexG = CadFlexG & "<" & String$(Len(CadCamp) - Len(NomCamp), " ") & "|"
'''       End Select
'''   Next i
'''  End With
''' ' MsgBox CadFlexG
'''  DBGMalla.FormatString = MidStrg(CadFlexG, 1, Len(CadFlexG) - 1)
'''  'MsgBox DBGMalla.FormatString
'''  DBGMalla.Visible = True
'''  RatonNormal
'''  DBGMalla.SetFocus
'''Exit Sub
'''Errorhandler:
'''    RatonNormal
'''    MsgBox ("Error #" & CStr(Err.Number) & " " & Err.Description)
'''    Err.Clear    ' Limpia el error.
'''    Exit Sub
'''End Sub

Public Sub ConectarAdodc(AdoBase As Adodc)
    AdoBase.ConnectionString = AdoStrCnn
End Sub

Public Sub ConectarAdodc_MySQL(AdoBase As Adodc)
    AdoBase.ConnectionString = AdoStrCnnMySQL
End Sub

'''Public Sub DesConectar_Adodc(AdoBase As Adodc)
'''  If AdoBase.Recordset.State = adStateOpen Then AdoBase.Recordset.Close
'''End Sub

Public Sub ConectarAdodcBackup(AdoBase As Adodc)
  AdoBase.ConnectionString = AdoStrCnnBackup
End Sub

Public Sub ConectarAdoRecordSet(SQLQuery As String)
' Cadena de Conección a la base de datos
  ' MsgBox SQLQuery
  Set AdoReg = New ADODB.Recordset
  AdoReg.CursorType = adOpenStatic
  AdoReg.CursorLocation = adUseClient
  AdoReg.open SQLQuery, AdoStrCnn, , , adCmdText
End Sub

Public Sub Select_Data_DBF(SQLQuery As String)
'''Dim Conexion_DBF As String
'''
'''Set Dato_DBF.Base_Datos = New ADODB.Connection
'''Set Dato_DBF.Registo = New ADODB.Recordset
'''
''''Cadena de Conección a la base de datos
''' Conexion_DBF = "Provider=Microsoft.Jet.OLEDB.4.0;" _
'''              & "Data Source=E:\VFP98\SISSALES;" _
'''              & "Extended Properties=dBASE IV;" _
'''              & "User ID=Admin;Password=;"
'''
''' MsgBox Conexion_DBF & vbCrLf & vbCrLf & SQLQuery
''' Dato_DBF.Base_Datos.Open Conexion_DBF
'''
'''
''' SQLQuery = CompilarSQL(SQLQuery)
''' Dato_DBF.Registo.Open SQLQuery, Dato_DBF.Base_Datos, , , adCmdText
'''
    'Dato_DBF.Carpeta = "E:\VFP98\SISSALES"
    If Len(Dato_DBF.Carpeta) > 1 And Len(SQLQuery) > 1 Then
       RatonReloj
       If Existe_Carpeta(Dato_DBF.Carpeta) Then
          Set Dato_DBF.Base_Datos = OpenDatabase(Dato_DBF.Carpeta, False, False, "FoxPro 3.0;")
          Set Dato_DBF.Registo = Dato_DBF.Base_Datos.OpenRecordset(SQLQuery)
       End If
      'MsgBox Dato_DBF.Base_Datos.Recordsets.Count
       RatonNormal
    End If
End Sub

Public Sub Close_DBF()
    Dato_DBF.Registo.Close
    Dato_DBF.Base_Datos.Close
End Sub

Public Sub Select_AdoDB(AdoReg As ADODB.Recordset, _
                        SQLQuery As String, _
                        Optional NombreFile As String)
 'MsgBox SQLQuery
  RatonReloj
  SQLQuery = CompilarSQL(SQLQuery)
  Generar_File_SQL NombreFile, SQLQuery
  Set AdoReg = New ADODB.Recordset
  AdoReg.CursorType = adOpenStatic   'adOpenDynamic
  AdoReg.CursorLocation = adUseClient
  AdoReg.LockType = adLockOptimistic
 'MsgBox SQLQuery
  AdoReg.open SQLQuery, AdoStrCnn, , , adCmdText
 'Determinamos el ancho de los campos
  CantCampos = AdoReg.fields.Count
  RatonNormal
End Sub

Public Sub Select_AdoDB_MySQL(AdoReg As ADODB.Recordset, _
                              SQLQuery As String)
'''   strBaseDatos = "diskcover_empresas"
'''   strServidor = "db.diskcoversystem.com"
'''   strUsuario = "diskcover"
'''   strPassword = "disk2017Cover"
'''   strPuerto = "13306"
'''   AdoStrCnnMySQL = "DRIVER={MySQL ODBC 3.51 Driver};" _
'''                  & "SERVER=" & strServidor & ";" _
'''                  & "DATABASE=" & strBaseDatos & ";" _
'''                  & "UID=" & strUsuario & ";" _
'''                  & "PASSWORD=" & strPassword & ";" _
'''                  & "PORT=" & strPuerto & ";"
                              
 'MsgBox SQLQuery & vbCrLf & AdoStrCnnMySQL
  RatonReloj
  SQLQuery = CompilarSQL(SQLQuery)
  Set AdoReg = New ADODB.Recordset
  AdoReg.CursorType = adOpenStatic   'adOpenDynamic
  AdoReg.CursorLocation = adUseClient
  AdoReg.LockType = adLockOptimistic
  AdoReg.open SQLQuery, AdoStrCnnMySQL, , , adCmdText
  
 'Determinamos el ancho de los campos
  CantCampos = AdoReg.fields.Count
 'MsgBox SQLQuery & vbCrLf & AdoStrCnnMySQL
  RatonNormal
 'MsgBox CantCampos
End Sub

Public Sub Select_AdoDBTabla(AdoReg As ADODB.Recordset, _
                            SQLQuery As String)
' Cadena de Conección a la base de datos
  RatonReloj
  Set AdoReg = New ADODB.Recordset
  AdoReg.CursorType = adOpenDynamic
  AdoReg.CursorLocation = adUseClient
' Consultamos las cuentas de la tabla
  SQLQuery = CompilarSQL(SQLQuery)
  AdoReg.open SQLQuery, AdoStrCnn, , , adCmdTable
  RatonNormal
End Sub

Public Sub Conectar_Ado_Execute_MySQL(SQLQuery As String, _
                                      Optional RegSN As Boolean)
Dim AdoCon1 As ADODB.Connection
Dim IdTime As Long
     RatonReloj
    'Consultamos las cuentas de la tabla
     SQLQuery = CompilarSQL(SQLQuery)
    'MsgBox SQLQuery & vbCrLf & String(70, "_") & vbCrLf & AdoStrCnnMySQL
     Set AdoCon1 = New ADODB.Connection
     If Ping_PC("db.diskcoversystem.com") Then
        AdoCon1.open AdoStrCnnMySQL
        AdoCon1.Execute SQLQuery, RegAfectados, adCmdText
        AdoCon1.Close
     End If
     RatonNormal
     If RegSN Then MsgBox "Registros Afectados: " & Format$(RegAfectados, "#,##0")
End Sub

Public Sub Conectar_Ado_ExecuteBackup(SQLQuery As String, _
                                    Optional RegSN As Boolean)
Dim AdoCon1 As ADODB.Connection
Dim IdTime As Long
  RatonReloj
 'Consultamos las cuentas de la tabla
  SQLQuery = CompilarSQL(SQLQuery)
 'MsgBox SQLQuery
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnnBackup
  AdoCon1.Execute SQLQuery, RegAfectados, adCmdText
  AdoCon1.Close
  RatonNormal
  If RegSN Then MsgBox "Registros Afectados: " & Format$(RegAfectados, "#,##0")
End Sub

Public Sub ConectarDataExecute(SQLQuery As String, _
                               Optional RegSN As Boolean)
Dim CnVFP As New ADODB.Connection
Dim rs As ADODB.Recordset
    If Len(Dato_DBF.Carpeta) > 1 Then
       RatonReloj
      'MsgBox SQLQuery
       Set rs = New ADODB.Recordset
       Set CnVFP = New ADODB.Connection
       CnVFP.CursorLocation = adUseClient
       'MsgBox Dato_DBF.Tipo_Base
       Select Case Dato_DBF.Tipo_Base
         Case "FOXPRO"
              CnVFP.open "Driver={Microsoft Visual FoxPro Driver};" _
                       & "SourceType=DBF;" _
                       & "SourceDB=" & Dato_DBF.Carpeta & "\"
              Set rs = CnVFP.Execute(SQLQuery)
              CnVFP.Close
         Case "ACCESS"
       'rs.Close
       End Select
       RatonNormal
    End If
End Sub

Public Function GetState(intState As Integer) As String
   Select Case intState
     Case adStateClosed
          GetState = "adStateClosed"
     Case adStateOpen
          GetState = "adStateOpen"
   End Select
End Function

Public Sub ImprimirAdo(Datas As Adodc, _
                       FinDoc As Boolean, _
                       FormaImp As Byte, _
                       SizeLetra As Integer, _
                       Optional EsCampoCorto As Boolean)
On Error GoTo Errorhandler

Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then

RatonReloj
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontSize = SizeLetra
     Printer.FontName = TipoArialNarrow
     Do While Not .EOF
        'MsgBox Printer.FontName
        PrinterAllFields CantCampos, PosLinea, Datas, True, False
        PosLinea = PosLinea + 0.36
        If Cuadricula Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
           PosLinea = PosLinea + 0.05
        End If
        If PosLinea >= LimiteAlto Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
        End If
       .MoveNext
     Loop
End If
End With
Imprimir_Linea_H PosLinea, InicioX, Ancho(CantCampos), Negro, True
RatonNormal
MensajeEncabData = ""
If FinDoc Then Printer.EndDoc
Cuadricula = False
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
End If
End Sub

Public Sub ImprimirAdodc(Datas As Adodc, _
                         FormaImp As Byte, _
                         SizeLetra As Integer, _
                         Optional EsCampoCorto As Boolean, _
                         Optional PiePagina As String)
On Error GoTo Errorhandler
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
Orientacion_Pagina = FormaImp
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
InicioX = 0.5: InicioY = 0
DataAnchoCampos InicioX, Datas, SizeLetra, TipoArialNarrow, Orientacion_Pagina, EsCampoCorto
'''Cadena = "Ancho:" & vbCrLf
'''For I = 0 To CantCampos
'''    Cadena = Cadena & Ancho(I) & vbCrLf
'''Next I
'''MsgBox Cadena & AnchoPapel & ", Paginas = " & EnDosPaginas
Pagina = 1
'Iniciamos la impresion
Printer.FontBold = False
With Datas.Recordset
 If .RecordCount > 0 Then
    .MoveFirst
     EncabezadoData Datas
     Printer.FontName = TipoArialNarrow
     If Cuadricula Then
        If EnDosPaginas > 0 Then
           Imprimir_Linea_H PosLinea, Ancho(0), AnchoPapel, Gris
        Else
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
        End If
        PosLinea = PosLinea + 0.05
     End If
     Printer.FontSize = SizeLetra
     Do While Not .EOF
        If EnDosPaginas > 0 Then
           PrinterAllFields EnDosPaginas, PosLinea, Datas, True, False
        Else
           PrinterAllFields CantCampos, PosLinea, Datas, True, False
        End If
        PosLinea = PosLinea + 0.36
        If Cuadricula Then
           If EnDosPaginas > 0 Then
              Imprimir_Linea_H PosLinea, Ancho(0), AnchoPapel, Gris
           Else
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
           End If
           PosLinea = PosLinea + 0.05
        End If
        If PosLinea >= LimiteAlto Then
           If Cuadricula Then
              If EnDosPaginas > 0 Then
                 Imprimir_Linea_H PosLinea, Ancho(0), AnchoPapel, Gris
              Else
                 Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
              End If
              PosLinea = PosLinea + 0.05
           End If
           Printer.NewPage
           EncabezadoData Datas
           Printer.FontSize = SizeLetra
           Printer.FontName = TipoArialNarrow
        End If
       .MoveNext
     Loop
     If EnDosPaginas > 0 Then
        Imprimir_Linea_H PosLinea, Ancho(0), AnchoPapel, Negro, True
     Else
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
     End If
    'Imprimimos en la segunda pagina el restante de la impresion
     If EnDosPaginas > 0 Then
        Printer.NewPage
        EncabezadoData Datas, True
        If Cuadricula Then
           Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
           PosLinea = PosLinea + 0.05
        End If
        Printer.FontSize = SizeLetra
        Printer.FontBold = False
       .MoveFirst
        Do While Not .EOF
           Printer.FontName = TipoArialNarrow
           PrinterAllFields CantCampos, PosLinea, Datas, True, False, True
           PosLinea = PosLinea + 0.36
           If Cuadricula Then
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Gris
              PosLinea = PosLinea + 0.05
           End If
           If PosLinea >= LimiteAlto Then
              Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos)
              Printer.NewPage
              EncabezadoData Datas, True
              Printer.FontSize = SizeLetra
           End If
          .MoveNext
        Loop
       .MoveFirst
        Imprimir_Linea_H PosLinea, Ancho(0), Ancho(CantCampos), Negro, True
     End If
 End If
End With
PosLinea = PosLinea + 0.05
If PiePagina <> "" Then PrinterTexto Ancho(0), PosLinea, PiePagina
Cuadricula = False
MensajeEncabData = "": SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = "": SQLMsg4 = ""
RatonNormal
Printer.EndDoc
Cuadricula = False
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
Else
    RatonNormal
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

Public Sub SetAnchoCampos(PictPrint As Object, Optional EsCorto As Boolean)
  Determina_Ancho_Tipos EsCorto
  'determinaAnchoTipos
 'Determinamos el ancho de los tipo de campo que tiene VB
  If TypeOf PictPrint Is mjwPDF Then
     'Anchos por default
      WidthBoolean = CSng(PictPrint.PDFGetTextWidth(CadBoolean))
      WidthDate = CSng(PictPrint.PDFGetTextWidth(ULCase(CadDate)))
      WidthTime = CSng(PictPrint.PDFGetTextWidth(CadTime))
      WidthByte = CSng(PictPrint.PDFGetTextWidth(CadByte))
      WidthInteger = CSng(PictPrint.PDFGetTextWidth(CadInteger))
      WidthLong = CSng(PictPrint.PDFGetTextWidth(CadLong))
      WidthSingle = CSng(PictPrint.PDFGetTextWidth(CadSingle))
      WidthDouble = CSng(PictPrint.PDFGetTextWidth(CadDouble))
      WidthCurrency = CSng(PictPrint.PDFGetTextWidth(CadCurrency))
  ElseIf TypeOf PictPrint Is Printer Then
     'Determinamos el ancho de impresora
      WidthBoolean = PictPrint.TextWidth(CadBoolean)
      WidthDate = PictPrint.TextWidth(ULCase(CadDate))
      WidthTime = PictPrint.TextWidth(CadTime)
      WidthByte = PictPrint.TextWidth(CadByte)
      WidthInteger = PictPrint.TextWidth(CadInteger)
      WidthLong = PictPrint.TextWidth(CadLong)
      WidthSingle = PictPrint.TextWidth(CadSingle)
      WidthDouble = PictPrint.TextWidth(CadDouble)
      WidthCurrency = PictPrint.TextWidth(CadCurrency)
  ElseIf TypeOf PictPrint Is DataGrid Then
     'Determinamos el ancho de la Maya
      WidthBoolean = CSng(PictPrint.Parent.TextWidth(CadBoolean))
      WidthDate = CSng(PictPrint.Parent.TextWidth(ULCase(CadDate)))
      WidthTime = CSng(PictPrint.Parent.TextWidth(CadTime))
      WidthByte = CSng(PictPrint.Parent.TextWidth(CadByte))
      WidthInteger = CSng(PictPrint.Parent.TextWidth(CadInteger))
      WidthLong = CSng(PictPrint.Parent.TextWidth(CadLong))
      WidthSingle = CSng(PictPrint.Parent.TextWidth(CadSingle))
      WidthDouble = CSng(PictPrint.Parent.TextWidth(CadDouble))
      WidthCurrency = CSng(PictPrint.Parent.TextWidth(CadCurrency))
  Else
      WidthBoolean = CSng(Screen.ActiveForm.TextWidth(CadBoolean))
      WidthDate = CSng(Screen.ActiveForm.TextWidth(ULCase(CadDate)))
      WidthTime = CSng(Screen.ActiveForm.TextWidth(CadTime))
      WidthByte = CSng(Screen.ActiveForm.TextWidth(CadByte))
      WidthInteger = CSng(Screen.ActiveForm.TextWidth(CadInteger))
      WidthLong = CSng(Screen.ActiveForm.TextWidth(CadLong))
      WidthSingle = CSng(Screen.ActiveForm.TextWidth(CadSingle))
      WidthDouble = CSng(Screen.ActiveForm.TextWidth(CadDouble))
      WidthCurrency = CSng(Screen.ActiveForm.TextWidth(CadCurrency))
  End If
End Sub

Public Function Maximo_De(Tabla As String, Campo As String) As Long
Dim DataReg As ADODB.Recordset
Dim RegMaximo As Long
  RegMaximo = 0
  If Campo <> "ID" Then
     RegMaximo = 1
     Set DataReg = New ADODB.Recordset
     DataReg.CursorType = adOpenStatic
     DataReg.CursorLocation = adUseClient
     sSQL = "SELECT MAX(" & Campo & ") As Maximo " _
          & "FROM " & Tabla & " " _
          & "WHERE " & Campo & " <> 0 "
     DataReg.open sSQL, AdoStrCnn, , , adCmdText
     If DataReg.RecordCount > 0 Then If Not IsNull(DataReg.fields("Maximo")) Then RegMaximo = DataReg.fields("Maximo") + 1
     DataReg.Close
  End If
  Maximo_De = RegMaximo
End Function

Public Function Datos_Nacion(Rubro As String, Tipo_R As String, Optional C_Pais As String, Optional C_Prov As String) As Datos_Naciones
Dim DataReg As ADODB.Recordset
Dim sSQLDN As String
Dim DN1 As Datos_Naciones
  Set DataReg = New ADODB.Recordset
  DataReg.CursorType = adOpenStatic
  DataReg.CursorLocation = adUseClient
  With DN1
      .CCiudad = ""
      .Codigo = ""
      .CPais = ""
      .CProvincia = ""
      .CRegion = ""
      .Descripcion = ""
      .Tipo_Rubro = ""
      .Pais = ""
      .Provincia = ""
  End With
  
  sSQLDN = "SELECT * " _
         & "FROM Tabla_Naciones " _
         & "WHERE TR = '" & Tipo_R & "' " _
         & "AND MidStrg(Descripcion_Rubro,1," & Len(Rubro) & ") = '" & Rubro & "' "
  If C_Pais <> "" Then sSQLDN = sSQLDN & "AND CPais = '" & C_Pais & "' "
  If C_Prov <> "" Then sSQLDN = sSQLDN & "AND CProvincia = '" & C_Prov & "' "
  sSQLDN = CompilarSQL(sSQLDN)
  DataReg.open sSQLDN, AdoStrCnn, , , adCmdText
  With DataReg
   If .RecordCount > 0 Then
       DN1.CCiudad = .fields("CCiudad")
       DN1.Codigo = .fields("Codigo")
       DN1.CPais = .fields("CPais")
       DN1.CProvincia = .fields("CProvincia")
       DN1.CRegion = .fields("CRegion")
       DN1.Descripcion = .fields("Descripcion_Rubro")
       DN1.Tipo_Rubro = .fields("TR")
   End If
  End With
  DataReg.Close
  sSQLDN = "SELECT * " _
         & "FROM Tabla_Naciones "
  If Val(C_Pais) = 593 Then
     sSQLDN = sSQLDN & "WHERE CPais = '" & DN1.CPais & "' "
  Else
    sSQLDN = sSQLDN & "WHERE CPais = '" & C_Pais & "' "
  End If
  sSQLDN = sSQLDN & "AND TR = 'N' "
  DataReg.open sSQLDN, AdoStrCnn, , , adCmdText
  With DataReg
   If .RecordCount > 0 Then
       DN1.Pais = .fields("Descripcion_Rubro")
       DN1.CPais = .fields("CPais")
   End If
  End With
  DataReg.Close
  sSQLDN = "SELECT * " _
         & "FROM Tabla_Naciones " _
         & "WHERE CPais = '" & DN1.CPais & "' " _
         & "AND CProvincia = '" & DN1.CProvincia & "' " _
         & "AND TR = 'P' "
  DataReg.open sSQLDN, AdoStrCnn, , , adCmdText
  With DataReg
   If .RecordCount > 0 Then
       DN1.Provincia = .fields("Descripcion_Rubro")
   End If
  End With
  DataReg.Close
'''  MsgBox "<----------->" & vbCrLf _
'''       & DN1.CCiudad & vbCrLf _
'''       & DN1.Codigo & vbCrLf _
'''       & DN1.CPais & vbCrLf _
'''       & DN1.CProvincia & vbCrLf _
'''       & DN1.CRegion & vbCrLf _
'''       & DN1.Descripcion & vbCrLf _
'''       & DN1.Tipo_Rubro & vbCrLf _
'''       & DN1.Pais & vbCrLf _
'''       & DN1.Provincia & vbCrLf
  Datos_Nacion = DN1
End Function

'''Public Sub GenerarArchivoExcel(DtaAux As Adodc)
'''Dim NumFile As Long
'''Dim RutaGeneraFile As String
'''Dim CaptionOld As String
'''Dim ValorBool As String
'''Dim ValorTexto As String
'''Dim apexcel As Variant
'''Dim Porcentaje  As Single
'''RatonReloj
'''Contador = 0
'''Progreso_Barra.Incremento = 0
'''Progreso_Barra.Valor_Maximo = 100
'''Progreso_Barra.Mensaje_Box = "Generando Archivo en Excel..."
'''Progreso_Esperar
''''''   ' aplica formula
''''''     apexcel.cells(3, 4).formula = "=b3-c3"
''''''   ' hace una seleccion de celdas y pone bordes de color
''''''     apexcel.range("b3:d3").borders.Color = RGB(255, 0, 0)
'''With DtaAux.Recordset
''' If .RecordCount > 0 Then
'''     Progreso_Barra.Valor_Maximo = .RecordCount + 1
'''     J = 1
'''    .MoveFirst
'''     Set apexcel = CreateObject("Excel.application")
'''    'hace que excel se vea
'''     apexcel.Visible = False
'''    'agrega un nuevo libro
'''     apexcel.workbooks.Add
'''     Contador = Contador + 1
'''     For I = 0 To .Fields.Count - 1
'''         apexcel.cells(Contador, I + 1).formula = .Fields(I).Name
'''     Next I
'''     Do While Not .EOF
'''        Contador = Contador + 1
'''       'MiFrom.Caption = RutaGeneraFile & ": Registro No. " & Format$(Contador / .RecordCount, "00%")
'''        For I = 0 To .Fields.Count - 1
'''            Select Case .Fields(I).Type
'''              Case TadDate, TadDate1
'''                   If IsNull(.Fields(I)) Then
'''                      apexcel.cells(Contador, I + 1).formula = "'" & CStr(Format$(FechaSistema, FormatoFechas))
'''                   Else
'''                      apexcel.cells(Contador, I + 1).formula = "'" & CStr(Format$(.Fields(I), FormatoFechas))
'''                   End If
'''              Case TadByte, TadInteger, TadLong
'''                   If IsNull(.Fields(I)) Then
'''                      apexcel.cells(Contador, I + 1).formula = "0"
'''                   Else
'''                      apexcel.cells(Contador, I + 1).formula = Format$(.Fields(I), "##0")
'''                   End If
'''              Case TadSingle, TadDouble
'''                   If IsNull(.Fields(I)) Then
'''                      apexcel.cells(Contador, I + 1).formula = "0.00"
'''                   Else
'''                      apexcel.cells(Contador, I + 1).formula = Format$(.Fields(I), "#,##0.000000")
'''                   End If
'''              Case TadCurrency
'''                   If IsNull(.Fields(I)) Then
'''                      apexcel.cells(Contador, I + 1).formula = "0.00"
'''                   Else
'''                      apexcel.cells(Contador, I + 1).formula = Format$(.Fields(I), "#,##0.0000")
'''                   End If
'''              Case TadBoolean
'''                   If IsNull(.Fields(I)) Then
'''                      ValorBool = "No"
'''                   Else
'''                      If .Fields(I) = 0 Then ValorBool = "No" Else ValorBool = "Si"
'''                   End If
'''                   apexcel.cells(Contador, I + 1).formula = ValorBool
'''              Case Else
'''                   If IsNull(.Fields(I)) Then
'''                      ValorTexto = "Ninguno"
'''                   Else
'''                      If CStr(.Fields(I)) = Ninguno Then ValorTexto = "" Else ValorTexto = "'" & CStr(.Fields(I))
'''                   End If
'''                   apexcel.cells(Contador, I + 1).formula = ValorTexto
'''            End Select
'''        Next I
'''        Porcentaje = Contador / .RecordCount
'''        Progreso_Barra.Mensaje_Box = "Procesando"
'''        Progreso_Esperar
'''       .MoveNext
'''     Loop
'''    .MoveFirst
'''    'poner titulos
'''     Contador = Contador + 1
'''     apexcel.cells(Contador, 1).Font.Size = 8
'''     apexcel.cells(Contador, 1).formula = .RecordCount & " Registros"
'''     apexcel.Visible = True
''' End If
'''End With
'''
'''Set apexcel = Nothing
'''RatonNormal
'''Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
'''Progreso_Final
'''End Sub

Public Function Existe_Tabla(Nombre_Tabla As String) As Boolean
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim Hay_Tabla As Boolean
' Consultamos las cuentas de la tabla
  RatonReloj
  Hay_Tabla = False
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
  Do Until RstSchema.EOF
     If RstSchema!TABLE_NAME = Nombre_Tabla Then Hay_Tabla = True
     RstSchema.MoveNext
  Loop
  AdoCon1.Close
  Existe_Tabla = Hay_Tabla
  RatonNormal
End Function

Public Function Existe_Campo(Nombre_Tabla As String, NCampo As String) As Boolean
Dim AdoAuxDB As ADODB.Recordset
Dim ExisteC As Boolean
Dim C As Integer
    ExisteC = False
    sSQL = "SELECT * " _
         & "FROM " & Nombre_Tabla & " " _
         & "WHERE 1 = 0 "
    Select_AdoDB AdoAuxDB, sSQL
    For C = 0 To AdoAuxDB.fields.Count - 1
        If NCampo = AdoAuxDB.fields(C).Name Then ExisteC = True
    Next C
    AdoAuxDB.Close
    Existe_Campo = ExisteC
End Function

Public Sub Actualiza_Cursos()
  If Periodo_Contable = Ninguno Then
     If SQL_Server Then
        sSQL = "UPDATE Clientes_Matriculas " _
             & "SET Grupo_No = C.Grupo " _
             & "FROM Clientes_Matriculas As CM,Clientes As C "
     Else
        sSQL = "UPDATE Clientes_Matriculas As CM,Clientes As C " _
             & "SET CM.Grupo_No = C.Grupo "
     End If
     sSQL = sSQL _
          & "WHERE CM.Item = '" & NumEmpresa & "' " _
          & "AND CM.Periodo = '" & Periodo_Contable & "' " _
          & "AND CM.Item = '" & NumEmpresa & "' " _
          & "AND C.FA <> " & Val(adFalse) & " " _
          & "AND CM.Codigo = C.Codigo " _
          & "AND CM.Grupo_No <> C.Grupo "
     Ejecutar_SQL_SP sSQL
  End If
End Sub

Public Sub Actualiza_Email(TEmail As String, CodigoCli As String)
   If Len(TEmail) > 1 And Len(CodigoCli) > 1 Then
      sSQL = "UPDATE Clientes " _
           & "SET Email = '" & TEmail & "' " _
           & "WHERE Codigo = '" & CodigoCli & "' "
        Ejecutar_SQL_SP sSQL
   End If
End Sub

Public Function Obtener_Clave(NombreDelUsuario As String) As String
Dim AdoUsuario As ADODB.Recordset
Dim ClaveDelUsuario As String
   ClaveDelUsuario = Ninguno
   sSQL = "SELECT * " _
        & "FROM Accesos " _
        & "WHERE UCaseStrg(Usuario) = '" & UCaseStrg(NombreDelUsuario) & "' "
   Select_AdoDB AdoUsuario, sSQL
   If AdoUsuario.RecordCount > 0 Then ClaveDelUsuario = AdoUsuario.fields("Clave")
   AdoUsuario.Close
   Obtener_Clave = ClaveDelUsuario
End Function

Public Sub Fechas_Balances(DetalleFecha As String, MBoxFechaI As MaskEdBox, MBoxFechaF As MaskEdBox)
Dim AdoFecha As ADODB.Recordset
Dim sSQL1 As String
   If DetalleFecha <> Ninguno Then
      FechaValida MBoxFechaI
      FechaValida MBoxFechaF
      DetalleFecha = TrimStrg(MidStrg(DetalleFecha, 1, 20))
      sSQL = "SELECT * " _
           & "FROM Fechas_Balance " _
           & "WHERE Detalle = '" & DetalleFecha & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' "
      Select_AdoDB AdoFecha, sSQL
      With AdoFecha
       If .RecordCount > 0 Then
           MBoxFechaI = .fields("Fecha_Inicial")
           MBoxFechaF = .fields("Fecha_Final")
       Else
           sSQL1 = "INSERT INTO Fechas_Balance " _
                & "(Detalle,Item,Periodo,Fecha_Inicial,Fecha_Final,Cerrado) " _
                & "VALUES " _
                & "('" & DetalleFecha & "','" & NumEmpresa & "','" & Periodo_Contable & "'," _
                & "#" & BuscarFecha(MBoxFechaI.Text) & "#,#" & BuscarFecha(MBoxFechaF.Text) & "#,0) "
           Ejecutar_SQL_SP sSQL1
        End If
      End With
      AdoFecha.Close
   Else
      MBoxFechaI = FechaSistema
      MBoxFechaF = FechaSistema
   End If
End Sub

Public Sub Update_Fechas(DetalleFecha As String, MBoxFechaI As MaskEdBox, MBoxFechaF As MaskEdBox)
Dim sSQLU As String
   If DetalleFecha <> Ninguno Then
      sSQLU = "UPDATE Fechas_Balance " _
            & "SET Fecha_Inicial = #" & BuscarFecha(MBoxFechaI.Text) & "#, Fecha_Final = #" & BuscarFecha(MBoxFechaF.Text) & "# " _
            & "WHERE Detalle = '" & DetalleFecha & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "AND Item = '" & NumEmpresa & "' "
      Ejecutar_SQL_SP sSQLU
   End If
End Sub

Public Sub Cerrar_Periodo_Tablas(Nombre_Tabla As String, FechaFinCierre As String)
 If Len(FechaFinCierre) = 10 And Len(Nombre_Tabla) > 1 Then
    sSQL = "UPDATE " & Nombre_Tabla & " " _
         & "SET Periodo = '" & FechaFinCierre & "' " _
         & "WHERE Periodo = '.' " _
         & "AND Item = '" & NumEmpresa & "' "
    Ejecutar_SQL_SP sSQL
 End If
End Sub

Public Function T_Fields(Nombre_del_Campo As String) As Variant
Dim Idc_V As Long
Dim Tot_Fields As Variant
    Tot_Fields = Ninguno
    If LBound(Vect_Dec) <= UBound(Vect_Dec) Then
       For Idc_V = LBound(Vect_Dec) To UBound(Vect_Dec)
           If Vect_Dec(Idc_V).Campo = Nombre_del_Campo Then
              Tot_Fields = Vect_Dec(Idc_V).SumaTotal
           End If
       Next Idc_V
    End If
    T_Fields = Tot_Fields
End Function

Public Sub Crear_Cierre_Mes(Optional AnioNuevo As String)
Dim AdoFecha As ADODB.Recordset
Dim sSQL1 As String
Dim FechaIni As String
Dim FechaFin As String
Dim Detalle_Mes As String
Dim Anio As String
Dim AnioI As Integer
Dim AnioF As Integer
Dim Mes As String
Dim AI As Integer
Dim MesNo As Integer
Dim C As Boolean
    AnioI = Year(FechaSistema)
    AnioF = Year(FechaSistema)
    'MsgBox AnioNuevo
    If AnioNuevo = "" Then
       FechaFin = BuscarFecha("31/12/" & AnioF)
       If SQL_Server Then
          sSQL1 = "SELECT YEAR(Fecha) As Anio "
       Else
          sSQL1 = "SELECT DATEPART('yyyy',Fecha) As Anio "
       End If
       sSQL1 = sSQL1 _
             & "FROM Comprobantes " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND Fecha <= #" & FechaFin & "# " _
             & "AND T <> 'A' "
       If SQL_Server Then
          sSQL1 = sSQL1 _
                & "GROUP BY YEAR(Fecha) " _
                & "ORDER BY YEAR(Fecha) "
       Else
          sSQL1 = sSQL1 _
                & "GROUP BY DATEPART('yyyy',Fecha) " _
                & "ORDER BY DATEPART('yyyy',Fecha) "
       End If
       Select_AdoDB AdoFecha, sSQL1
       If AdoFecha.RecordCount > 0 Then
          AnioI = AdoFecha.fields("Anio")
          AdoFecha.MoveLast
          AnioF = AdoFecha.fields("Anio")
          If AnioI < 2000 Then AnioI = 2000
       End If
    Else
        If Len(AnioNuevo) = 4 Then
           AnioI = AnioNuevo
           AnioF = AnioNuevo
        End If
    End If
    sSQL1 = "DELETE * " _
          & "FROM Fechas_Balance " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Detalle <= '" & CStr(AnioI) & "' "
    Ejecutar_SQL_SP sSQL1
    
    For AI = AnioI To AnioF
        Anio = CStr(AI)
        For MesNo = 1 To 12
            FechaIni = "01/" & Format$(MesNo, "00") & "/" & Anio
            FechaFin = UltimoDiaMes(FechaIni)
            Mes = MesesLetras(MesNo)
            Detalle_Mes = Anio & " " & Mes
            If Periodo_Contable = Ninguno Then
               If MesNo >= Month(FechaSistema) And AI = Year(FechaSistema) Then C = False
            Else
               C = True
            End If
            sSQL1 = "SELECT * " _
                  & "FROM Fechas_Balance " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "AND Detalle = '" & Detalle_Mes & "' "
            Select_AdoDB AdoFecha, sSQL1
            If AdoFecha.RecordCount <= 0 Then
               SetAdoAddNew "Fechas_Balance"
               SetAdoFields "Item", NumEmpresa
               SetAdoFields "Periodo", Periodo_Contable
               SetAdoFields "Detalle", Detalle_Mes
               SetAdoFields "Fecha_Inicial", FechaIni
               SetAdoFields "Fecha_Final", FechaFin
               SetAdoFields "Cerrado", C
               SetAdoUpdate
            End If
        Next MesNo
    Next AI
    AdoFecha.Close
End Sub

Public Sub Imprimir_Formato_Propio(Tipo_Formato As String, Inicio_Xo As Single, Inicio_Yo As Single)
' No.  Formato
'======================================
'  0   Nombre del Formatos (Formulario)
'  1   Encabezados
'  2   Cuadros
'  3   Lineas
'  4   Circulos
'  5   Archivos
'  6   Grafico
'  7   Texto
'  8   Texto justificado
'  9   Fin del Formato (Cualquier cosa)
' 10   DATO_EMP
' 11   DATO_DIR
' 12   DATO_TEL
' 13   DATO_RUC
' 14   DATO_COM
' 15   DATO_PAI

'FORMATOS DE IMPRESION:
'=====================
'LINEA
'CUADRO
'CIRCULO
'TEXTO
'TEXTOS
'ARCHIVO
'GRAFICO
'DATO_EMP
'DATO_DIR
'DATO_TEL
'DATO_RUC
'DATO_COM

Dim AdoFormatoP As ADODB.Recordset
Dim FP_Vertical As Boolean

Dim SQLFormatoP As String
Dim FP_RGB As String
Dim FP_Texto As String
Dim FP_TextFile As String
Dim FP_PathDibujo As String

Dim FP_R As Integer
Dim FP_G As Integer
Dim FP_B As Integer

Dim FP_NumFile As Long
Dim FP_Color As Long

Dim FP_Radio As Single
Dim FP_Pos_Xo As Single
Dim FP_Pos_Yo As Single
Dim FP_Pos_Xf As Single
Dim FP_Pos_Yf As Single
Dim Inicio_X As Single
Dim Inicio_Y As Single
Dim FP_Tamanio As Single
Dim FP_PosL_Yo As Single
Dim FP_PosL_Yf As Single
Dim FP_CentrarTextoW As Single

On Error GoTo Errorhandler
   'Establecemos los inicios por default
    Inicio_X = Inicio_Xo
    Inicio_Y = Inicio_Yo
    If Inicio_X < 0 Then Inicio_X = 0
    If Inicio_Y < 0 Then Inicio_Y = 0
    If Tipo_Formato = "" Then Tipo_Formato = Ninguno
   'Establecemos Espacios y seteos de impresion
    SQLFormatoP = "SELECT * " _
                & "FROM Formato_Propio " _
                & "WHERE TP = '" & Tipo_Formato & "' " _
                & "AND Item = '000' " _
                & "ORDER BY Num,Tipo_Objeto,Codigo,Texto,Pos_Xo,Pos_Yo,Pos_Xf,Pos_Yf "
    Select_AdoDB AdoFormatoP, SQLFormatoP
   'MsgBox Tipo_Formato
    With AdoFormatoP
     'MsgBox "HOLA" & vbCrLf & .RecordCount
     If .RecordCount > 0 Then
         If Inicio_X > 0 And Inicio_Y > 0 Then
            RatonReloj
            Do While Not .EOF
               If .fields("Tipo_Letra") <> Ninguno Then
                   Printer.FontName = .fields("Tipo_Letra")
               Else
                   Printer.FontName = TipoArial ' TipoVerdana
               End If
               FP_Radio = .fields("Radio")
               FP_Pos_Xo = .fields("Pos_Xo")
               FP_Pos_Yo = .fields("Pos_Yo")
               FP_Pos_Xf = .fields("Pos_Xf")
               FP_Pos_Yf = .fields("Pos_Yf")
               If FP_Pos_Xo > 0 And FP_Pos_Yo > 0 Then
                 Select Case UCaseStrg(.fields("Color"))
                   Case "NEGRO": FP_Color = Negro
                   Case "AZUL": FP_Color = Azul
                   Case "VERDE": FP_Color = Verde
                   Case "AGUAMARINA": FP_Color = Aguamarina
                   Case "ROJO": FP_Color = Rojo
                   Case "FUCSIA": FP_Color = Fucsia
                   Case "AMARILLO": FP_Color = Amarillo
                   Case "BLANCO": FP_Color = Blanco
                   Case "GRIS": FP_Color = Gris
                   Case "AZUL_CLARO": FP_Color = Azul_Claro
                   Case "VERDE_CLARO": FP_Color = Verde_Claro
                   Case "MAGENTA": FP_Color = Magenta
                   Case "ROJO_CLARO": FP_Color = Rojo_Claro
                   Case "FUCSIA_CLARO": FP_Color = Fucsia_Claro
                   Case "AMARILLO_CLARO": FP_Color = Amarillo_Claro
                   Case "BLANCO_BRILLANTE": FP_Color = Blanco
                   Case "BLANCO": FP_Color = Blanco
                   Case Else
                        FP_RGB = .fields("Color")
                        FP_R = CInt(SinEspaciosIzq(FP_RGB))
                        FP_RGB = TrimStrg(MidStrg(FP_RGB, Len(SinEspaciosIzq(FP_RGB)) + 1, Len(FP_RGB)))
                        FP_G = CInt(SinEspaciosIzq(FP_RGB))
                        FP_RGB = TrimStrg(MidStrg(FP_RGB, Len(SinEspaciosIzq(FP_RGB)) + 1, Len(FP_RGB)))
                        FP_B = CInt(TrimStrg(SinEspaciosIzq(FP_RGB)))
                        If (FP_R + FP_G + FP_B) = 0 Then FP_Color = Blanco Else FP_Color = RGB(FP_R, FP_G, FP_B)
                 End Select
                 If .fields("Negrilla") Then Printer.FontBold = True Else Printer.FontBold = False
                 Select Case .fields("Tipo_Objeto")
                   Case "DATO_EMP": FP_Texto = Empresa
                   Case "DATO_COM": FP_Texto = NombreComercial
                   Case "DATO_DIR": FP_Texto = Direccion & " * Telef. " & Telefono1 & "*"
                   Case "DATO_TEL": FP_Texto = Telefono1
                   Case "DATO_RUC": FP_Texto = "R.U.C. " & RUC
                   Case "DATO_PAI": FP_Texto = NombreCiudad & " - " & NombrePais
                   Case Else: FP_Texto = .fields("Texto")
                 End Select
                 FP_Vertical = .fields("Vertical")
                 FP_Tamanio = .fields("Tamaño")
                 If FP_Tamanio <= 0 Then FP_Tamanio = 1
                 FP_CentrarTextoW = FP_Pos_Xf - FP_Pos_Xo
                 If FP_CentrarTextoW < 0 Then FP_CentrarTextoW = 0
                 If FP_CentrarTextoW > Printer.TextWidth(FP_Texto) Then
                    FP_CentrarTextoW = FP_Pos_Xo + (FP_CentrarTextoW / 2) - (Printer.TextWidth(FP_Texto) / 2)
                 Else
                    FP_CentrarTextoW = FP_Pos_Xo
                 End If
                'Imprimir el Tipo de Objeto
                 Select Case .fields("Tipo_Objeto")
                   Case "LINEA"
                        Printer.DrawWidth = FP_Tamanio
                        Printer.Line (Inicio_X + FP_Pos_Xo, Inicio_Y + FP_Pos_Yo) _
                                    -(Inicio_X + FP_Pos_Xf, Inicio_Y + FP_Pos_Yf), FP_Color, B
                   Case "CUADRO"
                        Printer.DrawWidth = FP_Tamanio
                        Printer.Line (Inicio_X + FP_Pos_Xo, Inicio_Y + FP_Pos_Yo) _
                                    -(Inicio_X + FP_Pos_Xf, Inicio_Y + FP_Pos_Yf), FP_Color, BF
                                    
                        Printer.Line (Inicio_X + FP_Pos_Xo, Inicio_Y + FP_Pos_Yo) _
                                    -(Inicio_X + FP_Pos_Xf, Inicio_Y + FP_Pos_Yf), Negro, B
                   Case "CIRCULO"
                   
                   Case "TEXTO"
                        If FP_Texto = Ninguno Then FP_Texto = " "
                        Printer.FontSize = FP_Tamanio
                        Printer.ForeColor = FP_Color
                        If .fields("Centrar") Then
                            Printer.CurrentX = Inicio_X + FP_CentrarTextoW
                        Else
                            Printer.CurrentX = Inicio_X + FP_Pos_Xo
                        End If
                        Printer.CurrentY = Inicio_Y + FP_Pos_Yo
                        Printer.Print FP_Texto
                   Case "TEXTOS"
                        Printer.FontSize = FP_Tamanio
                        Printer.ForeColor = FP_Color
                        PosLinea = Printer_Texto_Justifica(Inicio_X + FP_Pos_Xo, _
                                                 Inicio_X + FP_Pos_Xf, Inicio_Y + FP_Pos_Yo, FP_Texto)
                   Case "ARCHIVO"
                        If FP_Texto <> Ninguno Then
                           Printer.ForeColor = FP_Color
                           FP_PathDibujo = RutaSistema & "\DOCUMENT\"
                           If Existe_File(FP_PathDibujo & FP_Texto & ".txt") Then
                              FP_PosL_Yo = Inicio_Y + FP_Pos_Yo
                              PosLinea = FP_PosL_Yo + 0.05
                              Printer.FontSize = FP_Tamanio
                              FP_NumFile = FreeFile
                              Open FP_PathDibujo & FP_Texto & ".txt" For Input As #FP_NumFile
                                Do While Not EOF(FP_NumFile)
                                   Line Input #FP_NumFile, FP_TextFile
                                   If FP_TextFile = "===" Then
                                      PosLinea = PosLinea + 0.05
                                      Printer.Line (Inicio_X + FP_Pos_Xo, PosLinea) _
                                                  -(Inicio_X + FP_Pos_Xf, PosLinea), Negro, B
                                      PosLinea = PosLinea + 0.05
                                   Else
                                      PosLinea = Printer_Texto_Justifica(Inicio_X + FP_Pos_Xo + 0.1, _
                                                 Inicio_X + FP_Pos_Xf - 0.1, _
                                                 PosLinea, FP_TextFile) - Printer.TextHeight("H")
                                   End If
                                Loop
                              Close #FP_NumFile
                              Printer.DrawWidth = 6
                              Printer.Line (Inicio_X + FP_Pos_Xo, FP_PosL_Yo) _
                                          -(Inicio_X + FP_Pos_Xf, PosLinea + 0.1), Negro, B
                           End If
                        End If
                   Case "GRAFICO"
                        FP_PathDibujo = Ninguno
                        If FP_Texto <> Ninguno Then
                           FP_PathDibujo = RutaSistema & "\LOGOS\"
                           If Existe_File(FP_PathDibujo & FP_Texto & ".gif") Then
                              FP_PathDibujo = RutaSistema & "\LOGOS\" & FP_Texto & ".gif"
                           ElseIf Existe_File(FP_PathDibujo & FP_Texto & ".jpg") Then
                              FP_PathDibujo = RutaSistema & "\LOGOS\" & FP_Texto & ".jpg"
                           End If
                        End If
                        If FP_PathDibujo <> Ninguno Then
                           Printer.PaintPicture LoadPicture(FP_PathDibujo), _
                                                Inicio_X + FP_Pos_Xo, Inicio_Y + FP_Pos_Yo, _
                                                FP_Pos_Xf - FP_Pos_Xo, FP_Pos_Yf - FP_Pos_Yo
                        End If
                 End Select
               End If
              .MoveNext
            Loop
         End If
     End If
    End With
    AdoFormatoP.Close
    RatonNormal
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub

Public Function Grabar_Archivo_Excel(AdoArchivo As Adodc, NombreArchivo As String) As String
Dim NRegistro As Integer
Dim RutaGeneraFile As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim Porcentaje As Single

  RatonReloj
  Progreso_Iniciar
 'Start a new workbook in Excel
  Set oExcel = CreateObject("Excel.Application")
  Set oBook = oExcel.Workbooks.Add
 'Add data to cells of the first worksheet in the new workbook
  Set oSheet = oBook.Worksheets(1)
  RutaGeneraFile = NombreArchivo
  RutaGeneraFile = Replace(RutaGeneraFile, ".", "")
  RutaGeneraFile = Replace(RutaGeneraFile, "/", "-")
  RutaGeneraFile = UCaseStrg(TrimStrg(RutaGeneraFile))
  Contador = 1
  With AdoArchivo.Recordset
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = .RecordCount
       RutaGeneraFile = RutaSysBases & "\Auditoria\" & RutaGeneraFile & ".xls"
       If Dir(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
       For NRegistro = 0 To .fields.Count - 1
           oSheet.Range(Chr(65 + NRegistro) & "1").value = .fields(NRegistro).Name
       Next NRegistro
       Do While Not .EOF
          Contador = Contador + 1
          For NRegistro = 0 To .fields.Count - 1
              oSheet.Range(Chr(65 + NRegistro) & CStr(Contador)).value = CStr(.fields(NRegistro))
          Next NRegistro
          Porcentaje = Contador / .RecordCount
          Progreso_Esperar
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
 'Save the Workbook and Quit Excel
  oBook.SaveAs RutaGeneraFile
  oExcel.Quit
  Progreso_Final
  RatonNormal
  Grabar_Archivo_Excel = RutaGeneraFile
End Function

Public Function Grabar_Archivo_ExcelDB(AdoArchivo As ADODB.Recordset, NombreArchivo As String) As String
Dim NRegistro As Integer
Dim RutaGeneraFile As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim Porcentaje As Single

  RatonReloj
  Progreso_Esperar
  Progreso_Barra.Incremento = 0
 'Start a new workbook in Excel
  Set oExcel = CreateObject("Excel.Application")
  Set oBook = oExcel.Workbooks.Add
 'Add data to cells of the first worksheet in the new workbook
  Set oSheet = oBook.Worksheets(1)
  RutaGeneraFile = NombreArchivo
  RutaGeneraFile = Replace(RutaGeneraFile, ".", "")
  RutaGeneraFile = Replace(RutaGeneraFile, "/", "-")
  RutaGeneraFile = UCaseStrg(TrimStrg(RutaGeneraFile))
  Contador = 1
  RutaSysBases = LeftStrg(CurDir$, 2) & "\SYSBASES"
  
  With AdoArchivo
  'MsgBox .Source
   If .RecordCount > 0 Then
       Progreso_Barra.Valor_Maximo = .RecordCount
       RutaGeneraFile = RutaSysBases & "\Auditoria\" & RutaGeneraFile & ".xls"
       If Dir(RutaGeneraFile) <> "" Then Kill RutaGeneraFile
       For NRegistro = 0 To .fields.Count - 1
           oSheet.Range(Chr(65 + NRegistro) & "1").value = .fields(NRegistro).Name
       Next NRegistro
       Do While Not .EOF
          Contador = Contador + 1
          For NRegistro = 0 To .fields.Count - 1
              oSheet.Range(Chr(65 + NRegistro) & CStr(Contador)).value = CStr(.fields(NRegistro))
          Next NRegistro
          Porcentaje = Contador / .RecordCount
          Progreso_Esperar
''          If (Contador And 1) = 0 Then
''             FEsperar.Label1.Caption = "Procesando el " & Format$(Porcentaje, "00%") & vbCrLf _
''                                     & String$(Contador Mod 10, ".") & "Espere por favor" & String$((Contador Mod 10) + 1, ".")
''          Else
''             FEsperar.Label1.Caption = vbCrLf _
''                                     & String$(Contador Mod 10, ".") & "Espere por favor" & String$((Contador Mod 10) + 1, ".")
''          End If
''          FEsperar.Label1.Refresh
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
 'Save the Workbook and Quit Excel
  oBook.SaveAs RutaGeneraFile
  oExcel.Quit
  Progreso_Final
  RatonNormal
  Grabar_Archivo_ExcelDB = RutaGeneraFile
End Function

Public Sub UPD_Abrir_Campos_SQL(NumFile As Integer)
    Cod_Emp = "": Cod_Base = "": Cod_Field = ""
    Line Input #NumFile, Cod_Base
    TotalReg = CLng(SinEspaciosIzq(Cod_Base))
    Cod_Base = SinEspaciosDer(Cod_Base)
    Line Input #NumFile, Cod_Field
    'MsgBox Cod_Base & vbCrLf & Cod_Field
    CantCamposUpdate = 0
    For I = 1 To Len(Cod_Field)
        If MidStrg(Cod_Field, I, 1) = "|" Then CantCamposUpdate = CantCamposUpdate + 1
    Next I
    ReDim TipoC(CantCamposUpdate + 1) As Campos_Tabla
    No_Desde = 1: No_Hasta = 1
    Cadena = Cod_Field
    For I = 1 To CantCamposUpdate
        Do
           No_Hasta = No_Hasta + 1
        Loop Until MidStrg(Cadena, No_Hasta, 1) = "|"
        TipoC(I).Campo = TrimStrg(MidStrg(Cadena, No_Desde, No_Hasta - 1))
        Cadena = MidStrg(Cadena, No_Hasta + 1, Len(Cadena))
        No_Desde = 1: No_Hasta = 1
    Next I
End Sub

Public Sub UPD_Actualizar_La_Tabla(ROrigen As String, Optional EsAccesos As Boolean)
Dim AdoCompDB As ADODB.Recordset
Dim AdoCompDB1 As ADODB.Recordset
Dim Crear_Campo As Boolean
  sSQL = "SELECT * " _
       & "FROM " & ROrigen & " "
  Select_AdoDB AdoCompDB, sSQL
  With AdoCompDB
   SQL1 = "CREATE TABLE A_" & ROrigen & " ("
   For K = 0 To .fields.Count - 1
       SQL1 = SQL1 & "[" & .fields(K).Name & "] "
       If SQL_Server Then
          Cadena1 = FieldTypeSQL(.fields(K).Type)
       Else
          Cadena1 = FieldTypeAccess(.fields(K).Type)
       End If
       SQL1 = SQL1 & Cadena1
       If Cadena1 = "NVARCHAR" Or Cadena1 = "TEXT" Then
          SQL1 = SQL1 & "(" & .fields(K).DefinedSize & ")"
       End If
       SQL1 = SQL1 & " NULL "
       If K <> (.fields.Count - 1) Then SQL1 = SQL1 & ","
   Next K
   SQL1 = SQL1 & "); "
   Ejecutar_SQL_SP SQL1
   SQL2 = "INSERT INTO A_" & ROrigen & " " _
        & "SELECT * " _
        & "FROM " & ROrigen & ";"
   Ejecutar_SQL_SP SQL2
 ' Eliminamos la Tabla actual
   SQL1 = "DROP TABLE [" & ROrigen & "] "
   Ejecutar_SQL_SP SQL1
 ' Volvemos a crear la tabla
   SQL1 = "CREATE TABLE " & ROrigen & " ("
   For K = 0 To .fields.Count - 1
       Crear_Campo = True
       If ROrigen = "Accesos" And .fields(K).Name = "Item" Then Crear_Campo = False
       If ROrigen = "Accesos" And .fields(K).Name = "CI" Then Crear_Campo = False
       If Crear_Campo Then
           SQL1 = SQL1 & "[" & .fields(K).Name & "] "
           If SQL_Server Then
              Cadena1 = FieldTypeSQL(.fields(K).Type)
           Else
              Cadena1 = FieldTypeAccess(.fields(K).Type)
           End If
           SQL1 = SQL1 & Cadena1
           If Cadena1 = "NVARCHAR" Or Cadena1 = "TEXT" Then
              If EsAccesos Then
                 If .fields(K).Name = "Codigo" Then
                     SQL1 = SQL1 & "(10)"
                 Else
                     SQL1 = SQL1 & "(" & .fields(K).DefinedSize & ")"
                 End If
              Else
                 If .fields(K).Name = "CodigoU" Or .fields(K).Name = "CodigoA" Then
                     SQL1 = SQL1 & "(10)"
                 Else
                     SQL1 = SQL1 & "(" & .fields(K).DefinedSize & ")"
                 End If
              End If
           End If
           SQL1 = SQL1 & " NULL,"
           'If K <> (.Fields.Count - 1) Then SQL1 = SQL1 & ","
       End If
   Next K
  'MsgBox SQL1
   SQL1 = MidStrg(SQL1, 1, Len(SQL1) - 1)
  'MsgBox SQL1
   SQL1 = SQL1 & "); "
   Ejecutar_SQL_SP SQL1
   SQL2 = "INSERT INTO " & ROrigen & " " _
        & "SELECT "
        
   sSQL = "SELECT * " _
        & "FROM " & ROrigen & " "
   Select_AdoDB AdoCompDB1, sSQL
   With AdoCompDB1
        SQL2 = "INSERT INTO " & ROrigen & " ("
        For K = 0 To .fields.Count - 1
            SQL2 = SQL2 & "[" & .fields(K).Name & "] "
            If K <> (.fields.Count - 1) Then SQL2 = SQL2 & ","
        Next K
        SQL2 = SQL2 & ") " _
             & "SELECT "
        For K = 0 To .fields.Count - 1
            SQL2 = SQL2 & "[" & .fields(K).Name & "] "
            If K <> (.fields.Count - 1) Then SQL2 = SQL2 & ","
        Next K
        SQL2 = SQL2 & "FROM A_" & ROrigen & ";"
   End With
   AdoCompDB1.Close
  'MsgBox SQL2
   Ejecutar_SQL_SP SQL2
 ' Eliminamos la Tabla temporal
   SQL1 = "DROP TABLE [A_" & ROrigen & "] "
   Ejecutar_SQL_SP SQL1
  End With
  AdoCompDB.Close
End Sub

Public Sub UPD_Leer_Campos_Tabla(RutaDatosTabla As String)
Dim NumFile As Long
Dim IdIni As Long
Dim IdFin As Long
   'MsgBox RutaDatosTabla
    NumFile = FreeFile
    Open RutaDatosTabla For Input As #NumFile
    Line Input #NumFile, Cod_Base
    CantCampos = CInt(TrimStrg(Cod_Base))
    ReDim TablaNew(CantCampos) As Crear_Tablas
    I = 0
    Do While Not EOF(NumFile)
       Line Input #NumFile, Cod_Field
       No_Desde = 1: No_Hasta = 1
       Cadena = Cod_Field
       TablaNew(I).ErrorCampo = False
       TablaNew(I).LargoCampo = 0
       For J = 1 To 3
           Do
             No_Hasta = No_Hasta + 1
           Loop Until MidStrg(Cadena, No_Hasta, 1) = "|"
           Select Case J
             Case 1: TablaNew(I).Campo = TrimStrg(MidStrg(Cadena, No_Desde, No_Hasta - 1))
             Case 2: TablaNew(I).TipoSQL = TrimStrg(MidStrg(Cadena, No_Desde, No_Hasta - 1))
             Case 3: TablaNew(I).TipoAccess = TrimStrg(MidStrg(Cadena, No_Desde, No_Hasta - 1))
           End Select
           Cadena = MidStrg(Cadena, No_Hasta + 1, Len(Cadena))
           No_Desde = 1: No_Hasta = 1
       Next J
       
       If SQL_Server Then
          IdIni = InStr(TablaNew(I).TipoSQL, "(")
          IdFin = InStr(TablaNew(I).TipoSQL, ")")
          If IdIni > 0 And IdFin > 0 Then TablaNew(I).LargoCampo = Val(MidStrg(TablaNew(I).TipoSQL, IdIni + 1, IdFin - IdIni))
       Else
          IdIni = InStr(TablaNew(I).TipoSQL, "(")
          IdFin = InStr(TablaNew(I).TipoSQL, ")")
          If IdIni > 0 And IdFin > 0 Then TablaNew(I).LargoCampo = Val(MidStrg(TablaNew(I).TipoAccess, IdIni + 1, IdFin - IdIni))
       End If
      'MsgBox RutaDatosTabla & vbCrLf & TablaNew(I).Campo & " -> " & TablaNew(I).LargoCampo
       I = I + 1
    Loop
    Close #NumFile
End Sub

Public Sub UPD_Actualizar_Tablas_Temporales(LstStatud As ListBox)
Dim IdFile As Long
Dim ContadorReg As Long
Dim NumFile As Integer
Dim Archivo As String
Dim NombreArchivo() As String
Dim RutaArchivo() As String
Dim ArchivoTexto As String
  'MsgBox strIPServidor
   RatonReloj
   ContadorReg = 0
  'Determinar cuales son las tablas fijas que se van a actualizar
  'Redim Preserve NombreArchivo(ContadorReg) As String
   sSQL = "DELETE * " _
        & "FROM Trans_Documentos " _
        & "WHERE Item = '000' "
   Ejecutar_SQL_SP sSQL
   Archivo = Dir(RutaSistema & "\BASES\UPDATE_DB\*.DBS", vbNormal) 'Recupera la primera entrada.
   Do While Archivo <> ""
      If Archivo <> "." And Archivo <> ".." Then
         ReDim Preserve RutaArchivo(ContadorReg) As String
         ReDim Preserve NombreArchivo(ContadorReg) As String
         RutaArchivo(ContadorReg) = RutaSistema & "\BASES\UPDATE_DB\" & Archivo
         NombreArchivo(ContadorReg) = TrimStrg(MidStrg(Archivo, 1, Len(Archivo) - 4))
         ContadorReg = ContadorReg + 1
      End If
      Archivo = Dir
   Loop
   
  'Subimos el contenidos de cada tabla de la actualizacion actual
   For IdFile = 0 To UBound(NombreArchivo)
''       Progreso_Barra.Mensaje_Box = "Subiendo: " & NombreArchivo(IdFile) & ", a la base de datos"
''       Progreso_Esperar True
       LstStatud.AddItem "Subiendo DB: " & RutaArchivo(IdFile)
       LstStatud.Text = "Subiendo DB: " & RutaArchivo(IdFile)
       LstStatud.Refresh
       
       ArchivoTexto = Leer_Archivo_Texto(RutaArchivo(IdFile))
      'If NombreArchivo(IdFile) = "ZTipo_Concepto_Retencion" Then MsgBox ArchivoTexto
       SetAdoAddNew "Trans_Documentos"
       SetAdoFields "Item", "000"
       SetAdoFields "Periodo", Ninguno
       SetAdoFields "Clave_Acceso", NombreArchivo(IdFile)
       SetAdoFields "Documento_Autorizado", ArchivoTexto
       SetAdoFields "TD", "UD"
       SetAdoFields "Serie", "000000"
       SetAdoFields "Documento", 0
       SetAdoUpdate
   Next IdFile
   RatonNormal
End Sub

Public Sub UPD_Actualizar_Datos_Defecto(ProgressBarEstado As ProgressBar, _
                                        LstStatud As ListBox, _
                                        URLInet As Inet, _
                                        Update_Dir As DirListBox, _
                                        Update_File As FileListBox, _
                                        Update_LstTablas As ListBox, _
                                        Optional Update_Limpiar_Bases As Boolean)
Dim AdoAuxDB As ADODB.Recordset
Dim AdoCompDB As ADODB.Recordset
Dim nCampos() As String
Dim Si_actualizo As Boolean
Dim Idx As Integer
Dim Idy As Integer
Dim Idz As Long
Dim CantCampos As Integer
 '========================================================================
 'Actualizaciones Extras sobre la nueva actualizacion
 '========================================================================
 'MsgBox "Presione Aceptar para empezar la actualizacion: " & Periodo_Contable
  RatonReloj
  UPD_Listar_Tablas Update_LstTablas
  
'''  Progreso_Barra.Mensaje_Box = "PROGRESO DEL ACTUALIZACION"
'''  Progreso_Iniciar
'''  Progreso_Barra.Valor_Maximo = (Update_LstTablas.ListCount * 5)
'''  Progreso_Esperar
 
  ConSubDir = False
  Contador = 0: FileResp = 0
  FechaInicial = "01/01/" & Year(FechaSistema)
  FechaFinal = UltimoDiaMes(FechaSistema)
  sSQL = "SELECT MIN(Fecha) As Fecha_MIN " _
       & "FROM Facturas " _
       & "WHERE Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_AdoDB AdoCompDB, sSQL
  If AdoCompDB.RecordCount > 0 Then
     If Not IsNull(AdoCompDB.fields("Fecha_MIN")) Then FechaInicial = PrimerDiaMes(AdoCompDB.fields("Fecha_MIN"))
  End If
  AdoCompDB.Close

  sSQL = "SELECT MIN(Fecha) As Fecha_MIN " _
       & "FROM Trans_Abonos " _
       & "WHERE Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Select_AdoDB AdoCompDB, sSQL
  If AdoCompDB.RecordCount > 0 Then
     If Not IsNull(AdoCompDB.fields("Fecha_MIN")) Then
        If CFechaLong(FechaInicial) > CFechaLong(AdoCompDB.fields("Fecha_MIN")) Then
           FechaInicial = PrimerDiaMes(AdoCompDB.fields("Fecha_MIN"))
        End If
     End If
  End If
  AdoCompDB.Close

 'Volvemos a actualizar las tablas actuales despues de haber borrado las temporales o vacias
  UPD_Listar_Tablas Update_LstTablas
 'Actualizando datos, inserciones y eliminaciones de las tablas de esta actualizacion
  Mifecha = "31/12/" & Format(Year(FechaSistema) + 1, "0000")
  For I = 0 To Update_LstTablas.ListCount - 1
      Progreso_Barra.Mensaje_Box = "Actualizando datos por defecto de " & Update_LstTablas.List(I)
      ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
      sSQL = "SELECT * " _
           & "FROM " & Update_LstTablas.List(I) & " " _
           & "WHERE 1 = 0 "
      Select_AdoDB AdoAuxDB, sSQL
      ReDim nCampos(AdoAuxDB.fields.Count) As String
      For K = 0 To UBound(nCampos) - 1
          nCampos(K) = AdoAuxDB.fields(K).Name
      Next K
      AdoAuxDB.Close
      
     'MsgBox "Empezamos la actualizacion"
      Select Case Update_LstTablas.List(I)
        Case "Catalogo_Rol_Cuentas"
             Update_Default_SP "Catalogo_Rol_Cuentas"
        Case "Catalogo_Rol_Pagos"
             Update_Default_SP "Catalogo_Rol_Pagos"
        Case "Clientes"
             Update_Default_SP "Clientes"
        Case "Clientes_Matriculas"
'             sSQL = "UPDATE Clientes_Matriculas " _
'                  & "SET Cedula_R = REPLACE(Cedula_R,'´','') " _
'                  & "WHERE Item <> '-' " _
'                  & "AND Periodo = '" & Periodo_Contable & "' "
'             Ejecutar_SQL_SP sSQL
'
'             sSQL = "UPDATE Clientes_Matriculas " _
'                  & "SET TD = 'R', Cedula_R = '9999999999999' " _
'                  & "WHERE Representante = 'CONSUMIDOR FINAL' " _
'                  & "AND Periodo = '" & Periodo_Contable & "' " _
'                  & "AND TD <> 'R' "
'             Ejecutar_SQL_SP sSQL
'
'             sSQL = "UPDATE Clientes_Matriculas " _
'                  & "SET TD = 'R', Representante = 'CONSUMIDOR FINAL' " _
'                  & "WHERE Cedula_R = '9999999999999' " _
'                  & "AND TD <> 'R' " _
'                  & "AND Periodo = '" & Periodo_Contable & "' "
'             Ejecutar_SQL_SP sSQL
             
             sSQL = "SELECT Periodo, Item, Codigo, Cedula_R, TD, Representante " _
                  & "FROM Clientes_Matriculas " _
                  & "WHERE TD = '" & Ninguno & "' " _
                  & "AND Periodo = '" & Periodo_Contable & "' " _
                  & "ORDER BY Representante, Periodo, Item, Codigo "
             Select_AdoDB AdoAuxDB, sSQL
             Contador = 0
             If AdoAuxDB.RecordCount > 0 Then
                Do While Not AdoAuxDB.EOF
                   Contador = Contador + 1
                   SubCtaGen = Format(Contador / AdoAuxDB.RecordCount, "00%")
'''                   Progreso_Barra.Mensaje_Box = "(" & SubCtaGen & ") Representante: " & AdoAuxDB.Fields("Periodo") & " " & AdoAuxDB.Fields("Representante")
'''                   Progreso_Esperar True
                   DigVerif = Digito_Verificador(AdoAuxDB.fields("Cedula_R"))
                   AdoAuxDB.fields("TD") = Tipo_RUC_CI.Tipo_Beneficiario
                   AdoAuxDB.Update
                   AdoAuxDB.MoveNext
                Loop
             End If
             AdoAuxDB.Close
'''        Case "Codigosxx"
'''            'Numeracion de Retenciones
'''             sSQL = "SELECT Periodo, Item, Serie_Retencion, MAX(SecRetencion) As Ret_No " _
'''                  & "FROM Trans_Compras " _
'''                  & "WHERE LEN(Serie_Retencion)=6 " _
'''                  & "GROUP BY Periodo, Item, Serie_Retencion " _
'''                  & "ORDER BY Periodo, Item, Serie_Retencion "
'''             Select_AdoDB AdoAuxDB, sSQL
'''             Contador = 0
'''             If AdoAuxDB.RecordCount > 0 Then
'''                Do While Not AdoAuxDB.EOF
'''                   Contador = Contador + 1
'''                   SubCtaGen = Format(Contador / AdoAuxDB.RecordCount, "00%")
''''''                   Progreso_Barra.Mensaje_Box = "(" & SubCtaGen & ") Ret. Serie: " & AdoAuxDB.Fields("Serie_Retencion")
''''''                   Progreso_Esperar True
'''                   sSQL = "SELECT * " _
'''                        & "FROM Codigos " _
'''                        & "WHERE Periodo = '" & AdoAuxDB.Fields("Periodo") & "' " _
'''                        & "AND Item = '" & AdoAuxDB.Fields("Item") & "' " _
'''                        & "AND Concepto = 'RE_SERIE_" & AdoAuxDB.Fields("Serie_Retencion") & "' "
'''                   Select_AdoDB AdoCompDB, sSQL
'''                  'MsgBox "REt Retencion: " & AdoCompDB.RecordCount
'''                   If AdoCompDB.RecordCount > 0 Then
'''                      AdoCompDB.Fields("Numero") = AdoAuxDB.Fields("Ret_No") + 1
'''                      AdoCompDB.Update
'''                   Else
'''                      SetAdoAddNew "Codigos"
'''                      SetAdoFields "Item", AdoAuxDB.Fields("Item")
'''                      SetAdoFields "Periodo", AdoAuxDB.Fields("Periodo")
'''                      SetAdoFields "Concepto", "RE_SERIE_" & AdoAuxDB.Fields("Serie_Retencion")
'''                      SetAdoFields "Numero", AdoAuxDB.Fields("Ret_No") + 1
'''                      SetAdoUpdate
'''                   End If
'''                   AdoCompDB.Close
'''                   AdoAuxDB.MoveNext
'''                Loop
'''             End If
'''             AdoAuxDB.Close
'''        Case "Detalle_Facturaxx"
'''             sSQL = "UPDATE Detalle_Factura " _
'''                  & "SET Serie = '001001', Autorizacion = '0123456789' " _
'''                  & "WHERE Serie = '.' " _
'''                  & "AND Periodo = '" & Periodo_Contable & "' "
'''             Ejecutar_SQL_SP sSQL
'''             For Idz = CFechaLong(FechaInicial) To CFechaLong(FechaFinal) Step 31
'''                 Idx = Month(CLongFecha(Idz))
'''                 Idy = Year(CLongFecha(Idz))
''''''                 Progreso_Barra.Mensaje_Box = "Actualizando datos por defecto de Detalle_Factura (R) => " & Idx & "/" & Idy
''''''                 Progreso_Esperar False
'''                'Temporalmente hasta que esten actualizados todos
'''                 If Existe_Campo("Trans_Abonos", "EstabRetencion") Then
'''                    sSQL = "UPDATE Trans_Abonos " _
'''                         & "SET EstabRetencion = '001' " _
'''                         & "WHERE MidStrg(Banco,1,16) = 'RETENCION FUENTE' " _
'''                         & "AND EstabRetencion IN ('.','000') " _
'''                         & "AND MONTH(Fecha) = " & Idx & " " _
'''                         & "AND YEAR(Fecha) = " & Idy & " " _
'''                         & "AND Periodo = '" & Periodo_Contable & "' "
'''                    Ejecutar_SQL_SP sSQL
'''
'''                    sSQL = "UPDATE Trans_Abonos " _
'''                         & "SET PtoEmiRetencion = '001' " _
'''                         & "WHERE MidStrg(Banco,1,16) = 'RETENCION FUENTE' " _
'''                         & "AND PtoEmiRetencion IN ('.','000') " _
'''                         & "AND MONTH(Fecha) = " & Idx & " " _
'''                         & "AND YEAR(Fecha) = " & Idy & " " _
'''                         & "AND Periodo = '" & Periodo_Contable & "' "
'''                    Ejecutar_SQL_SP sSQL
'''
'''                    sSQL = "UPDATE Trans_Abonos " _
'''                         & "SET Serie_R = EstabRetencion+PtoEmiRetencion " _
'''                         & "WHERE Serie_R = '.' " _
'''                         & "AND MidStrg(Banco,1,16) = 'RETENCION FUENTE' " _
'''                         & "AND MONTH(Fecha) = " & Idx & " " _
'''                         & "AND YEAR(Fecha) = " & Idy & " " _
'''                         & "AND Periodo = '" & Periodo_Contable & "' "
'''                    Ejecutar_SQL_SP sSQL
'''                 End If
'''                 sSQL = "UPDATE Trans_Abonos " _
'''                      & "SET Secuencial_R = CAST(Cheque AS int) " _
'''                      & "WHERE Secuencial_R = 0 " _
'''                      & "AND MidStrg(Banco,1,16) = 'RETENCION FUENTE' " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND Secuencial_R = 0 "
'''                 Ejecutar_SQL_SP sSQL
'''
''''''                 Progreso_Barra.Mensaje_Box = "Actualizando datos por defecto de Detalle_Factura => " & Idx & "/" & Idy
''''''                 Progreso_Esperar False
'''                 sSQL = "UPDATE Detalle_Factura " _
'''                      & "SET Mes = '" & MesesLetras(Idx) & "', Mes_No = " & Idx & " " _
'''                      & "WHERE Mes = '" & MesesLetras(Idx) & "' " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Mes_No = 0 " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 sSQL = "UPDATE Detalle_Factura " _
'''                      & "SET Mes = '" & MesesLetras(Idx) & "', Mes_No = " & Idx & " " _
'''                      & "WHERE Mes = '.' " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 If SQL_Server Then
'''                    sSQL = "UPDATE Detalle_Factura " _
'''                         & "SET Cta_Venta = CP.Cta_Ventas " _
'''                         & "FROM Detalle_Factura As DF,Catalogo_Productos As CP "
'''                 Else
'''                    sSQL = "UPDATE Detalle_Factura As DF,Catalogo_Productos As CP " _
'''                         & "SET DF.Cta_Venta = CP.Cta_Ventas "
'''                 End If
'''                 sSQL = sSQL _
'''                      & "WHERE MONTH(DF.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(DF.Fecha) = " & Idy & " " _
'''                      & "AND DF.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND LEN(DF.Cta_Venta) <= 1 " _
'''                      & "AND DF.Total_IVA <> 0 " _
'''                      & "AND DF.Item = CP.Item " _
'''                      & "AND DF.Periodo = CP.Periodo " _
'''                      & "AND DF.Codigo = CP.Codigo_Inv "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 If SQL_Server Then
'''                    sSQL = "UPDATE Detalle_Factura " _
'''                         & "SET Cta_Venta = CP.Cta_Ventas_0 " _
'''                         & "FROM Detalle_Factura As DF,Catalogo_Productos As CP "
'''                 Else
'''                    sSQL = "UPDATE Detalle_Factura As DF,Catalogo_Productos As CP " _
'''                         & "SET DF.Cta_Venta = CP.Cta_Ventas_0 "
'''                 End If
'''                 sSQL = sSQL _
'''                      & "WHERE MONTH(DF.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(DF.Fecha) = " & Idy & " " _
'''                      & "AND DF.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND LEN(DF.Cta_Venta) <= 1 " _
'''                      & "AND DF.Total_IVA = 0 " _
'''                      & "AND DF.Item = CP.Item " _
'''                      & "AND DF.Periodo = CP.Periodo " _
'''                      & "AND DF.Codigo = CP.Codigo_Inv "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 If SQL_Server Then
'''                    sSQL = "UPDATE Detalle_Factura " _
'''                         & "SET Fecha_NC = TA.Fecha " _
'''                         & "FROM Detalle_Factura AS DF,Trans_Abonos AS TA "
'''                 Else
'''                    sSQL = "UPDATE Detalle_Factura AS DF,Trans_Abonos AS TA " _
'''                         & "SET DF.Fecha_NC = TA.Fecha "
'''                 End If
'''                 sSQL = sSQL _
'''                      & "WHERE TA.Banco ='NOTA DE CREDITO' " _
'''                      & "AND MONTH(DF.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(DF.Fecha) = " & Idy & " " _
'''                      & "AND DF.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND DF.TC = TA.TP " _
'''                      & "AND DF.Serie = TA.Serie " _
'''                      & "AND DF.Autorizacion = TA.Autorizacion " _
'''                      & "AND DF.Factura = TA.Factura " _
'''                      & "AND DF.Periodo = TA.Periodo " _
'''                      & "AND DF.Item = TA.Item " _
'''                      & "AND DF.CodigoC = TA.CodigoC "
'''                 Ejecutar_SQL_SP sSQL
'''             Next Idz
        Case "Empresas"
             Update_Default_SP "Empresas"
'''        Case "Facturasxx"
'''             'MsgBox "Inicio ->"
'''
'''             sSQL = "UPDATE Facturas " _
'''                  & "SET X = '.' " _
'''                  & "WHERE Item <> '.' "
'''             Ejecutar_SQL_SP sSQL
'''
'''             If SQL_Server Then
'''                sSQL = "UPDATE Facturas " _
'''                     & "SET X = 'X' " _
'''                     & "FROM Facturas As F,Facturas_Formatos As TF "
'''             Else
'''                sSQL = "UPDATE Facturas As F,Facturas_Formatos As TF " _
'''                     & "SET F.X = 'X' "
'''             End If
'''             sSQL = sSQL _
'''                  & "WHERE F.Fecha BETWEEN TF.Fecha_Inicio AND TF.Fecha_Final " _
'''                  & "AND F.Item = TF.Item " _
'''                  & "AND F.Periodo = TF.Periodo " _
'''                  & "AND F.TC = TF.TC " _
'''                  & "AND F.Serie = TF.Serie " _
'''                  & "AND F.Autorizacion = TF.Autorizacion "
'''             Ejecutar_SQL_SP sSQL
'''
'''             sSQL = "INSERT INTO Facturas_Formatos (Cod_CxC, Periodo, Item, TC, Serie, Autorizacion, Fecha_Inicio, Fecha_Final) " _
'''                  & "SELECT Cod_CxC, Periodo, Item, TC, Serie, Autorizacion, MIN(Fecha) As FMin, MAX(Fecha) As FMax " _
'''                  & "FROM Facturas " _
'''                  & "WHERE X = '.' " _
'''                  & "AND LEN(Autorizacion) <= 13 " _
'''                  & "GROUP BY Cod_CxC, Periodo, Item, TC, Serie, Autorizacion "
'''             Ejecutar_SQL_SP sSQL
'''
'''             sSQL = "UPDATE Catalogo_Lineas " _
'''                  & "SET X = '.' " _
'''                  & "WHERE Item <> '.' "
'''             Ejecutar_SQL_SP sSQL
'''
'''             If SQL_Server Then
'''                sSQL = "UPDATE Catalogo_Lineas " _
'''                     & "SET X = 'X' " _
'''                     & "FROM Catalogo_Lineas As CL, Facturas_Formatos As F "
'''             Else
'''                sSQL = "UPDATE Catalogo_Lineas As CL, Facturas_Formatos As F " _
'''                     & "SET CL.X = 'X' "
'''             End If
'''             sSQL = sSQL _
'''                  & "WHERE LEN(CL.Autorizacion) = 13 " _
'''                  & "AND CL.Item = F.Item " _
'''                  & "AND CL.Periodo = F.Periodo " _
'''                  & "AND CL.Fact = F.TC " _
'''                  & "AND CL.Serie = F.Serie " _
'''                  & "AND CL.Autorizacion = F.Autorizacion "
'''             Ejecutar_SQL_SP sSQL
'''
'''             sSQL = "INSERT INTO Facturas_Formatos (Cod_CxC, Periodo, Item, TC, Serie, Autorizacion, Fecha_Inicio, Fecha_Final) " _
'''                  & "SELECT Codigo, Periodo, Item, Fact, Serie, Autorizacion, Fecha, Vencimiento " _
'''                  & "FROM Catalogo_Lineas " _
'''                  & "WHERE X = '.' " _
'''                  & "AND LEN(Autorizacion) = 13 "
'''             Ejecutar_SQL_SP sSQL
'''
'''             If SQL_Server Then
'''                sSQL = "UPDATE Facturas_Formatos " _
'''                     & "SET Formato_Factura = CL.Logo_Factura, Largo = CL.Largo, Ancho = CL.Ancho, Espacios = CL.Espacios, Pos_Factura = CL.Pos_Factura," _
'''                     & "Fact_Pag = CL.Fact_Pag, Pos_Y_Fact = CL.Pos_Y_Fact, Nombre_Establecimiento = CL.Nombre_Establecimiento," _
'''                     & "Direccion_Establecimiento = CL.Direccion_Establecimiento, Telefono_Estab = CL.Telefono_Estab, " _
'''                     & "Logo_Tipo_Estab=CL.Logo_Tipo_Estab, Tipo_Impresion=CL.Tipo_Impresion, Concepto=CL.Concepto " _
'''                     & "FROM Facturas_Formatos As FF, Catalogo_Lineas As CL "
'''             Else
'''                sSQL = "UPDATE Facturas_Formatos As FF, Catalogo_Lineas As CL " _
'''                     & "SET FF.Formato_Factura = CL.Logo_Factura, FF.Largo = CL.Largo, FF.Ancho = CL.Ancho, FF.Espacios = CL.Espacios, FF.Pos_Factura = CL.Pos_Factura," _
'''                     & "FF.Fact_Pag = CL.Fact_Pag, FF.Pos_Y_Fact = CL.Pos_Y_Fact, FF.Nombre_Establecimiento = CL.Nombre_Establecimiento," _
'''                     & "FF.Direccion_Establecimiento = CL.Direccion_Establecimiento, FF.Telefono_Estab = CL.Telefono_Estab, " _
'''                     & "FF.Logo_Tipo_Estab=CL.Logo_Tipo_Estab, FF.Tipo_Impresion=CL.Tipo_Impresion, FF.Concepto=CL.Concepto "
'''             End If
'''             sSQL = sSQL _
'''                  & "WHERE FF.Item = CL.Item " _
'''                  & "AND FF.Periodo = CL.Periodo " _
'''                  & "AND FF.Cod_CxC = CL.Codigo " _
'''                  & "AND FF.TC = CL.Fact " _
'''                  & "AND FF.Serie = CL.Serie "
'''             Ejecutar_SQL_SP sSQL
'''
'''             'MsgBox "-> FIN"
'''
'''             Mifecha = FechaSistema
'''             FechaFinal = UltimoDiaMes(FechaSistema)
'''             For Idz = CFechaLong(FechaInicial) To CFechaLong(FechaFinal) Step 31
'''                 Idx = Month(CLongFecha(Idz))
'''                 Idy = Year(CLongFecha(Idz))
'''''''                 Progreso_Barra.Mensaje_Box = "Actualizando datos por defecto de Facturas => " & Idx & "/" & Idy
'''''''                 Progreso_Esperar False
'''                 sSQL = "UPDATE Detalle_Factura " _
'''                      & "SET Total_Desc = ROUND(Total_Desc,2,0) " _
'''                      & "WHERE Item = '" & NumEmpresa & "' " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Total_Desc <> ROUND(Total_Desc,2,0) " _
'''                      & "AND Total_Desc > 0 "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 sSQL = "UPDATE Detalle_Factura " _
'''                      & "SET Total_Desc2 = ROUND(Total_Desc2,2,0) " _
'''                      & "WHERE Item = '" & NumEmpresa & "' " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Total_Desc2 <> ROUND(Total_Desc2,2,0) " _
'''                      & "AND Total_Desc2 > 0 "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                'Actualizamos Representante de la factura si no fuera para colegios
'''                 If SQL_Server Then
'''                    sSQL = "UPDATE Facturas " _
'''                         & "SET RUC_CI = C.CI_RUC, Razon_Social = C.Cliente, TB = C.TD, Direccion_RS = C.Direccion, " _
'''                         & "Telefono_RS = C.Telefono " _
'''                         & "FROM Facturas As F, Clientes As C "
'''                 Else
'''                    sSQL = "UPDATE Facturas As F, Clientes As C " _
'''                         & "SET F.RUC_CI = C.CI_RUC, F.Razon_Social = C.Cliente, F.TB = C.TD, F.Direccion_RS = C.Direccion, " _
'''                         & "F.Telefono_RS = C.Telefono "
'''                 End If
'''                 sSQL = sSQL _
'''                      & "WHERE LEN(F.Autorizacion) = 13 " _
'''                      & "AND C.TD IN ('C','R','P') " _
'''                      & "AND MONTH(F.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(F.Fecha) = " & Idy & " " _
'''                      & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND F.CodigoC = C.Codigo "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                'Actualizamos facturas si no tienen representante
'''                 If SQL_Server Then
'''                    sSQL = "UPDATE Facturas " _
'''                         & "SET RUC_CI = CM.Cedula_R, Razon_Social = CM.Representante, TB = CM.TD, " _
'''                         & "Direccion_RS = CM.Lugar_Trabajo_R, Telefono_RS = CM.Telefono_RS " _
'''                         & "FROM Facturas As F, Clientes_Matriculas As CM "
'''                 Else
'''                    sSQL = "UPDATE Facturas As F, Clientes_Matriculas As CM " _
'''                         & "SET F.RUC_CI = CM.Cedula_R, F.Razon_Social = CM.Representante, F.TB = CM.TD, " _
'''                         & "F.Direccion_RS = CM.Lugar_Trabajo_R, F.Telefono_RS = CM.Telefono_RS "
'''                 End If
'''                 sSQL = sSQL _
'''                      & "WHERE CM.TD IN ('C','R','P') " _
'''                      & "AND LEN(F.Razon_Social) <= 1 " _
'''                      & "AND LEN(CM.Representante) > 1 " _
'''                      & "AND MONTH(F.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(F.Fecha) = " & Idy & " " _
'''                      & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND F.Item = CM.Item " _
'''                      & "AND F.Periodo = CM.Periodo " _
'''                      & "AND F.CodigoC = CM.Codigo "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                'Actualizamos facturas con campos faltante
'''                 If SQL_Server Then
'''                    sSQL = "UPDATE Facturas " _
'''                         & "SET Direccion_RS = C.Direccion, Telefono_RS = C.Telefono " _
'''                         & "FROM Facturas As F, Clientes As C "
'''                 Else
'''                    sSQL = "UPDATE Facturas As F, Clientes As C " _
'''                         & "SET F.Direccion_RS = C.Direccion, F.Telefono_RS = C.Telefono "
'''                 End If
'''                 sSQL = sSQL _
'''                      & "WHERE LEN(F.Direccion_RS) <= 1 " _
'''                      & "AND MONTH(F.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(F.Fecha) = " & Idy & " " _
'''                      & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND F.CodigoC = C.Codigo "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 If SQL_Server Then
'''                    sSQL = "UPDATE Facturas " _
'''                         & "SET Direccion_RS = CM.Lugar_Trabajo_R, Telefono_RS = CM.Telefono_RS " _
'''                         & "FROM Facturas As F, Clientes_Matriculas As CM "
'''                 Else
'''                    sSQL = "UPDATE Facturas As F, Clientes_Matriculas As CM " _
'''                         & "SET F.Direccion_RS = CM.Lugar_Trabajo_R, F.Telefono_RS = CM.Telefono_RS "
'''                 End If
'''                 sSQL = sSQL _
'''                      & "WHERE LEN(CM.Lugar_Trabajo_R) > 1 " _
'''                      & "AND MONTH(F.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(F.Fecha) = " & Idy & " " _
'''                      & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND F.Item = CM.Item " _
'''                      & "AND F.Periodo = CM.Periodo " _
'''                      & "AND F.CodigoC = CM.Codigo "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                'Llenamos de CONSUMIDOR FINAL en las facturas con nulos
'''                 SQL1 = "UPDATE Facturas " _
'''                      & "SET RUC_CI='9999999999999', TB='R', Razon_Social='CONSUMIDOR FINAL' " _
'''                      & "WHERE LEN(RUC_CI) <= 1 " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "OR LEN(Razon_Social) <= 1 "
'''                 Ejecutar_SQL_SP SQL1
'''
'''                 sSQL = "UPDATE Facturas " _
'''                      & "SET Serie='001001', Autorizacion='0123456789', Vencimiento=#" & BuscarFecha(Mifecha) & "# " _
'''                      & "WHERE Serie = '.' " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' "
'''                 Ejecutar_SQL_SP sSQL
'''
''''''                 Progreso_Barra.Mensaje_Box = "Actualizando datos por defecto de Facturas: Descuentos de " & Idz & "/" & Idy
''''''                 Progreso_Esperar True
'''                 sSQL = "UPDATE Facturas " _
'''                      & "SET Desc_0 = (SELECT SUM(Total_Desc+Total_Desc2) " _
'''                      & "              FROM Detalle_Factura " _
'''                      & "              WHERE Detalle_Factura.Total_IVA = 0 " _
'''                      & "              AND MONTH(Detalle_Factura.Fecha) = " & Idx & " " _
'''                      & "              AND YEAR(Detalle_Factura.Fecha) = " & Idy & " " _
'''                      & "              AND Detalle_Factura.TC = Facturas.TC " _
'''                      & "              AND Detalle_Factura.Item = Facturas.Item " _
'''                      & "              AND Detalle_Factura.Periodo = Facturas.Periodo " _
'''                      & "              AND Detalle_Factura.Fecha = Facturas.Fecha " _
'''                      & "              AND Detalle_Factura.Factura = Facturas.Factura " _
'''                      & "              AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
'''                      & "              AND Detalle_Factura.Serie = Facturas.Serie " _
'''                      & "              AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
'''                      & "WHERE TC IN ('FA','NV','LC') " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND Desc_0 = 0 " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 sSQL = "UPDATE Facturas " _
'''                      & "SET Desc_X = (SELECT SUM(Total_Desc+Total_Desc2) " _
'''                      & "              FROM Detalle_Factura " _
'''                      & "              WHERE Detalle_Factura.Total_IVA > 0 " _
'''                      & "              AND MONTH(Detalle_Factura.Fecha) = " & Idx & " " _
'''                      & "              AND YEAR(Detalle_Factura.Fecha) = " & Idy & " " _
'''                      & "              AND Detalle_Factura.TC = Facturas.TC " _
'''                      & "              AND Detalle_Factura.Item = Facturas.Item " _
'''                      & "              AND Detalle_Factura.Periodo = Facturas.Periodo " _
'''                      & "              AND Detalle_Factura.Fecha = Facturas.Fecha " _
'''                      & "              AND Detalle_Factura.Factura = Facturas.Factura " _
'''                      & "              AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
'''                      & "              AND Detalle_Factura.Serie = Facturas.Serie " _
'''                      & "              AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
'''                      & "WHERE TC IN ('FA','NV','LC') " _
'''                      & "AND Desc_X = 0 " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 sSQL = "UPDATE Facturas " _
'''                      & "SET Descuento = (SELECT SUM(Total_Desc) " _
'''                      & "                 FROM Detalle_Factura " _
'''                      & "                 WHERE Detalle_Factura.Total_Desc > 0 " _
'''                      & "                 AND Detalle_Factura.TC = Facturas.TC " _
'''                      & "                 AND Detalle_Factura.Item = Facturas.Item " _
'''                      & "                 AND Detalle_Factura.Periodo = Facturas.Periodo " _
'''                      & "                 AND Detalle_Factura.Factura = Facturas.Factura " _
'''                      & "                 AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
'''                      & "                 AND Detalle_Factura.Serie = Facturas.Serie " _
'''                      & "                 AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
'''                      & "WHERE TC IN ('FA','NV','LC') " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Descuento = 0 "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 sSQL = "UPDATE Facturas " _
'''                      & "SET Descuento2 = (SELECT SUM(Total_Desc2) " _
'''                      & "                  FROM Detalle_Factura " _
'''                      & "                  WHERE Detalle_Factura.Total_Desc2 > 0 " _
'''                      & "                  AND Detalle_Factura.TC = Facturas.TC " _
'''                      & "                  AND Detalle_Factura.Item = Facturas.Item " _
'''                      & "                  AND Detalle_Factura.Periodo = Facturas.Periodo " _
'''                      & "                  AND Detalle_Factura.Factura = Facturas.Factura " _
'''                      & "                  AND Detalle_Factura.CodigoC = Facturas.CodigoC " _
'''                      & "                  AND Detalle_Factura.Serie = Facturas.Serie " _
'''                      & "                  AND Detalle_Factura.Autorizacion = Facturas.Autorizacion) " _
'''                      & "WHERE TC IN ('FA','NV','LC') " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Descuento2 = 0 "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                'Retenciones en la fuente
'''                 If SQL_Server Then
'''                    sSQL = "UPDATE Facturas " _
'''                         & "SET Fecha_R=TA.Fecha, Serie_R=TA.Serie_R, Secuencial_R=TA.Cheque, Autorizacion_R=TA.Comprobante " _
'''                         & "FROM Facturas As F,Trans_Abonos As TA "
'''                 Else
'''                    sSQL = "UPDATE Facturas As F,Trans_Abonos As TA " _
'''                         & "SET F.Fecha_R=TA.Fecha, F.Serie_R=TA.Serie_R, F.Secuencial_R=TA.Cheque, F.Autorizacion_R=TA.Comprobante "
'''                 End If
'''                 sSQL = sSQL _
'''                      & "WHERE MidStrg(TA.Banco,1,16) = 'RETENCION FUENTE' " _
'''                      & "AND MONTH(F.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(F.Fecha) = " & Idy & " " _
'''                      & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND F.Item = TA.Item " _
'''                      & "AND F.Periodo = TA.Periodo " _
'''                      & "AND F.TC = TA.TP " _
'''                      & "AND F.Serie = TA.Serie " _
'''                      & "AND F.Factura = TA.Factura " _
'''                      & "AND F.Autorizacion = TA.Autorizacion " _
'''                      & "AND F.CodigoC = TA.CodigoC "
'''                 Ejecutar_SQL_SP sSQL
'''             Next Idz
'''             sSQL = "UPDATE Facturas " _
'''                  & "SET Desc_0 = 0 " _
'''                  & "WHERE Desc_0 IS NULL "
'''             Ejecutar_SQL_SP sSQL
'''             sSQL = "UPDATE Facturas " _
'''                  & "SET Desc_X = 0 " _
'''                  & "WHERE Desc_X IS NULL "
'''             Ejecutar_SQL_SP sSQL
'''             sSQL = "UPDATE Facturas " _
'''                  & "SET Descuento = 0 " _
'''                  & "WHERE Descuento IS NULL "
'''             Ejecutar_SQL_SP sSQL
'''             sSQL = "UPDATE Facturas " _
'''                  & "SET Descuento2 = 0 " _
'''                  & "WHERE Descuento2 IS NULL "
'''             Ejecutar_SQL_SP sSQL
'''             sSQL = "UPDATE Facturas " _
'''                  & "SET Tipo_Pago = '01' " _
'''                  & "WHERE LEN(Tipo_Pago) <= 1 "
'''             Ejecutar_SQL_SP sSQL
'''             sSQL = "UPDATE Facturas " _
'''                  & "SET Porc_IVA = 0 " _
'''                  & "WHERE Porc_IVA IS NULL "
'''             Ejecutar_SQL_SP sSQL
'''
'''             sSQL = "SELECT Porc, Fecha_Inicio, Fecha_Final " _
'''                  & "FROM Tabla_Por_ICE_IVA " _
'''                  & "WHERE IVA <> " & Val(adFalse) & " " _
'''                  & "ORDER BY Fecha_Inicio, Fecha_Final,Porc DESC "
'''             Select_AdoDB AdoAuxDB, sSQL
'''             If AdoAuxDB.RecordCount > 0 Then
'''                Do While Not AdoAuxDB.EOF
'''                   FechaIni = AdoAuxDB.Fields("Fecha_Inicio")
'''                   FechaFin = AdoAuxDB.Fields("Fecha_Final")
'''                   Porc_IVA = AdoAuxDB.Fields("Porc") / 100
'''                   For Idy = Year(FechaInicial) To Year(FechaSistema)
'''                    For Idx = 1 To 12
''''''                        Progreso_Barra.Mensaje_Box = "Porc IVA " & Porc_IVA * 100 & " Fecha: " & Idx & "-" & Idy
''''''                        Progreso_Esperar True
'''                        sSQL = "UPDATE Facturas " _
'''                             & "SET Porc_IVA = ROUND(IVA/(Con_IVA-Desc_X),2,0) " _
'''                             & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaIni) & "# AND #" & BuscarFecha(FechaFin) & "# " _
'''                             & "AND MONTH(Fecha) = " & Idx & " " _
'''                             & "AND YEAR(Fecha) = " & Idy & " " _
'''                             & "AND Porc_IVA = 0 " _
'''                             & "AND Con_IVA > 0 "
'''                        Ejecutar_SQL_SP sSQL
'''
'''                        sSQL = "UPDATE Facturas " _
'''                             & "SET Porc_IVA = " & Porc_IVA & " " _
'''                             & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaIni) & "# AND #" & BuscarFecha(FechaFin) & "# " _
'''                             & "AND MONTH(Fecha) = " & Idx & " " _
'''                             & "AND YEAR(Fecha) = " & Idy & " " _
'''                             & "AND Porc_IVA = 0 "
'''                        Ejecutar_SQL_SP sSQL
'''
''''''                        sSQL = "UPDATE Facturas " _
''''''                             & "SET Porc_IVA = " & Porc_IVA & " " _
''''''                             & "WHERE Fecha BETWEEN #" & BuscarFecha(FechaIni) & "# AND #" & BuscarFecha(FechaFin) & "# " _
''''''                             & "AND MONTH(Fecha) = " & Idx & " " _
''''''                             & "AND YEAR(Fecha) = " & Idy & " " _
''''''                             & "AND Periodo = '" & Periodo_Contable & "' " _
''''''                             & "AND Porc_IVA BETWEEN -100 and " & Porc_IVA & " "
''''''                        Ejecutar_SQL_SP sSQL
'''                    Next Idx
'''                   Next Idy
'''                   AdoAuxDB.MoveNext
'''                Loop
'''             End If
'''             AdoAuxDB.Close
'''
'''             sSQL = "UPDATE Detalle_Factura " _
'''                  & "SET Porc_IVA = 0 " _
'''                  & "WHERE Porc_IVA IS NULL "
'''             Ejecutar_SQL_SP sSQL
'''             For Idy = Year(FechaInicial) To Year(FechaSistema)
'''                For Idx = 1 To 12
''''''                    Progreso_Barra.Mensaje_Box = "Detalle Factura, Fecha: " & Idx & "-" & Idy
''''''                    Progreso_Esperar True
'''                    If SQL_Server Then
'''                       sSQL = "UPDATE Detalle_Factura " _
'''                            & "SET Porc_IVA = F.Porc_IVA " _
'''                            & "FROM Detalle_Factura As DF,Facturas As F "
'''                    Else
'''                       sSQL = "UPDATE Detalle_Factura As DF,Facturas As F " _
'''                            & "SET DF.Porc_IVA = F.Porc_IVA "
'''                    End If
'''                    sSQL = sSQL _
'''                         & "WHERE DF.Porc_IVA = 0 " _
'''                         & "AND MONTH(DF.Fecha) = " & Idx & " " _
'''                         & "AND YEAR(DF.Fecha) = " & Idy & " " _
'''                         & "AND DF.Item = F.Item " _
'''                         & "AND DF.Periodo = F.Periodo " _
'''                         & "AND DF.TC = F.TC " _
'''                         & "AND DF.Serie = F.Serie " _
'''                         & "AND DF.Factura = F.Factura " _
'''                         & "AND DF.Autorizacion = F.Autorizacion "
'''                    Ejecutar_SQL_SP sSQL
'''                Next Idx
'''             Next Idy
'''
'''            'Numeracion de Facturas
'''             sSQL = "SELECT Periodo, Item, TC, Serie, MAX(Factura) As TC_No " _
'''                  & "FROM Facturas " _
'''                  & "WHERE LEN(TC) = 2 " _
'''                  & "AND LEN(Serie) = 6 " _
'''                  & "GROUP BY Periodo, Item, TC, Serie " _
'''                  & "ORDER BY Periodo, Item, TC, Serie "
'''             Select_AdoDB AdoAuxDB, sSQL
'''             Contador = 0
'''             If AdoAuxDB.RecordCount > 0 Then
'''                Do While Not AdoAuxDB.EOF
'''                   Contador = Contador + 1
'''                   SubCtaGen = Format(Contador / AdoAuxDB.RecordCount, "00%")
'''''                   Progreso_Barra.Mensaje_Box = "(" & SubCtaGen & ") TC. Serie: " & AdoAuxDB.Fields("Serie")
'''''                   Progreso_Esperar True
'''                   sSQL = "SELECT * " _
'''                        & "FROM Codigos " _
'''                        & "WHERE Periodo = '" & AdoAuxDB.Fields("Periodo") & "' " _
'''                        & "AND Item = '" & AdoAuxDB.Fields("Item") & "' " _
'''                        & "AND Concepto = '" & AdoAuxDB.Fields("TC") & "_SERIE_" & AdoAuxDB.Fields("Serie") & "' "
'''                   Select_AdoDB AdoCompDB, sSQL
'''                  'MsgBox "REt Retencion: " & AdoCompDB.RecordCount
'''                   If AdoCompDB.RecordCount > 0 Then
'''                      AdoCompDB.Fields("Numero") = AdoAuxDB.Fields("TC_No") + 1
'''                      AdoCompDB.Update
'''                   Else
'''                      SetAdoAddNew "Codigos"
'''                      SetAdoFields "Item", AdoAuxDB.Fields("Item")
'''                      SetAdoFields "Periodo", AdoAuxDB.Fields("Periodo")
'''                      SetAdoFields "Concepto", AdoAuxDB.Fields("TC") & "_SERIE_" & AdoAuxDB.Fields("Serie")
'''                      SetAdoFields "Numero", AdoAuxDB.Fields("TC_No") + 1
'''                      SetAdoUpdate
'''                   End If
'''                   AdoCompDB.Close
'''                   AdoAuxDB.MoveNext
'''                Loop
'''             End If
'''             AdoAuxDB.Close
'''        Case "Facturas_Auxiliaresxx"
'''            'Numeracion de Guias de Remision
'''             sSQL = "SELECT Periodo, Item, Serie_GR, MAX(Remision) As GR_No " _
'''                  & "FROM Facturas_Auxiliares " _
'''                  & "WHERE LEN(Serie_GR)=6 " _
'''                  & "GROUP BY Periodo, Item, Serie_GR " _
'''                  & "ORDER BY Periodo, Item, Serie_GR "
'''             Select_AdoDB AdoAuxDB, sSQL
'''             Contador = 0
'''             If AdoAuxDB.RecordCount > 0 Then
'''                Do While Not AdoAuxDB.EOF
'''                   Contador = Contador + 1
'''                   SubCtaGen = Format(Contador / AdoAuxDB.RecordCount, "00%")
''''''                   Progreso_Barra.Mensaje_Box = "(" & SubCtaGen & ") GR Serie: " & AdoAuxDB.Fields("Serie_GR")
''''''                   Progreso_Esperar True
'''                   sSQL = "SELECT * " _
'''                        & "FROM Codigos " _
'''                        & "WHERE Periodo = '" & AdoAuxDB.Fields("Periodo") & "' " _
'''                        & "AND Item = '" & AdoAuxDB.Fields("Item") & "' " _
'''                        & "AND Concepto = 'GR_SERIE_" & AdoAuxDB.Fields("Serie_GR") & "' "
'''                   Select_AdoDB AdoCompDB, sSQL
'''                  'MsgBox "REt Retencion: " & AdoCompDB.RecordCount
'''                   If AdoCompDB.RecordCount > 0 Then
'''                      AdoCompDB.Fields("Numero") = AdoAuxDB.Fields("GR_No") + 1
'''                      AdoCompDB.Update
'''                   Else
'''                      SetAdoAddNew "Codigos"
'''                      SetAdoFields "Item", AdoAuxDB.Fields("Item")
'''                      SetAdoFields "Periodo", AdoAuxDB.Fields("Periodo")
'''                      SetAdoFields "Concepto", "GR_SERIE_" & AdoAuxDB.Fields("Serie_GR")
'''                      SetAdoFields "Numero", AdoAuxDB.Fields("GR_No") + 1
'''                      SetAdoUpdate
'''                   End If
'''                   AdoCompDB.Close
'''                   AdoAuxDB.MoveNext
'''                Loop
'''             End If
'''             AdoAuxDB.Close
        Case "Seteos_Documentos"
             Update_Default_SP "Seteos_Documentos"
        Case "Ctas_Proceso"
            'Hasta que se actualicen todos
             Update_Default_SP "Ctas_Proceso"
'''        Case "Trans_Abonosxx"
''''''             Progreso_Barra.Mensaje_Box = "Borrando columnas de Trans_Abonos => " & Idx
''''''             Progreso_Esperar True
''''             sSQL = "ALTER TABLE Trans_Abonos " _
''''                  & "DROP COLUMN EstabRetencion;"
''''             Ejecutar_SQL_SP sSQL
'''             FechaFinal = UltimoDiaMes(FechaSistema)
'''             For Idz = CFechaLong(FechaInicial) To CFechaLong(FechaFinal) Step 31
'''                 Idx = Month(CLongFecha(Idz))
'''                 Idy = Year(CLongFecha(Idz))
''''''                 Progreso_Barra.Mensaje_Box = "Actualizando datos por defecto de Trans_Abonos => " & Idx
''''''                 Progreso_Esperar False
'''                 If Existe_Campo("Trans_Abonos", "EstabRetencion") Then
'''                    sSQL = "UPDATE Trans_Abonos " _
'''                         & "SET Serie_R=EstabRetencion+PtoEmiRetencion, Autorizacion_R=Comprobante " _
'''                         & "WHERE LEN(Serie_R)<=1 " _
'''                         & "AND LEN(EstabRetencion+PtoEmiRetencion)>2 " _
'''                         & "AND MONTH(Fecha) = " & Idx & " " _
'''                         & "AND YEAR(Fecha) = " & Idy & " " _
'''                         & "AND Periodo = '" & Periodo_Contable & "' "
'''                    Ejecutar_SQL_SP sSQL
'''                 End If
'''
'''                 sSQL = "UPDATE Trans_Abonos " _
'''                      & "SET Serie = '001001', Autorizacion = '0123456789' " _
'''                      & "WHERE Serie = '.' " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 sSQL = "UPDATE Facturas " _
'''                      & "SET Tipo_Pago = CC.Tipo_Pago " _
'''                      & "FROM Facturas AS F, Trans_Abonos AS TA, Catalogo_Cuentas AS CC " _
'''                      & "WHERE Len(CC.Tipo_Pago) >= 2 " _
'''                      & "AND F.Tipo_Pago <= '00' " _
'''                      & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND MONTH(F.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(F.Fecha) = " & Idy & " " _
'''                      & "AND F.Item = TA.Item " _
'''                      & "AND F.Periodo = TA.Periodo " _
'''                      & "AND TA.Item = CC.Item " _
'''                      & "AND TA.Periodo = CC.Periodo " _
'''                      & "AND TA.Cta = CC.Codigo " _
'''                      & "AND F.TC = TA.TP " _
'''                      & "AND F.Serie = TA.Serie " _
'''                      & "AND F.Autorizacion = TA.Autorizacion " _
'''                      & "AND F.Factura = TA.Factura " _
'''                      & "AND F.CodigoC = TA.CodigoC "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 If SQL_Server Then
'''                    sSQL = "UPDATE Trans_Abonos " _
'''                         & "SET Tipo_Cta = CC.TC " _
'''                         & "FROM Trans_Abonos As TA, Catalogo_Cuentas As CC "
'''                 Else
'''                    sSQL = "UPDATE Trans_Abonos As TA, Catalogo_Cuentas As CC " _
'''                         & "SET TA.Tipo_Cta = CC.TC "
'''                 End If
'''                 sSQL = sSQL _
'''                      & "WHERE TA.Item = '" & NumEmpresa & "' " _
'''                      & "AND TA.Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND MONTH(TA.Fecha) = " & Idx & " " _
'''                      & "AND YEAR(TA.Fecha) = " & Idy & " " _
'''                      & "AND TA.Item = CC.Item " _
'''                      & "AND TA.Periodo = CC.Periodo " _
'''                      & "AND TA.Cta = CC.Codigo "
'''                 Ejecutar_SQL_SP sSQL
'''
'''                 sSQL = "UPDATE Trans_Abonos " _
'''                      & "SET Abono = ROUND(Abono,2,0) " _
'''                      & "WHERE Item = '" & NumEmpresa & "' " _
'''                      & "AND Periodo = '" & Periodo_Contable & "' " _
'''                      & "AND MONTH(Fecha) = " & Idx & " " _
'''                      & "AND YEAR(Fecha) = " & Idy & " " _
'''                      & "AND Abono <> ROUND(Abono,2,0) "
'''                 Ejecutar_SQL_SP sSQL
'''             Next Idz
'''
'''            'Numeracion de Nota de Credito
'''             sSQL = "SELECT Periodo, Item, Serie_NC, MAX(Secuencial_NC) As NC_No " _
'''                  & "FROM Trans_Abonos " _
'''                  & "WHERE LEN(Serie_NC) = 6 " _
'''                  & "GROUP BY Periodo, Item, Serie_NC " _
'''                  & "ORDER BY Periodo, Item, Serie_NC "
'''             Select_AdoDB AdoAuxDB, sSQL
'''             Contador = 0
'''             If AdoAuxDB.RecordCount > 0 Then
'''                Do While Not AdoAuxDB.EOF
'''                   Contador = Contador + 1
'''                   SubCtaGen = Format(Contador / AdoAuxDB.RecordCount, "00%")
''''''                   Progreso_Barra.Mensaje_Box = "(" & SubCtaGen & ") NC Serie: " & AdoAuxDB.Fields("Serie_NC")
''''''                   Progreso_Esperar True
'''                   sSQL = "SELECT * " _
'''                        & "FROM Codigos " _
'''                        & "WHERE Periodo = '" & AdoAuxDB.Fields("Periodo") & "' " _
'''                        & "AND Item = '" & AdoAuxDB.Fields("Item") & "' " _
'''                        & "AND Concepto = 'NC_SERIE_" & AdoAuxDB.Fields("Serie_NC") & "' "
'''                   Select_AdoDB AdoCompDB, sSQL
'''                  'MsgBox "REt Retencion: " & AdoCompDB.RecordCount
'''                   If AdoCompDB.RecordCount > 0 Then
'''                      AdoCompDB.Fields("Numero") = AdoAuxDB.Fields("NC_No") + 1
'''                      AdoCompDB.Update
'''                   Else
'''                      SetAdoAddNew "Codigos"
'''                      SetAdoFields "Item", AdoAuxDB.Fields("Item")
'''                      SetAdoFields "Periodo", AdoAuxDB.Fields("Periodo")
'''                      SetAdoFields "Concepto", "NC_SERIE_" & AdoAuxDB.Fields("Serie_NC")
'''                      SetAdoFields "Numero", AdoAuxDB.Fields("NC_No") + 1
'''                      SetAdoUpdate
'''                   End If
'''                   AdoCompDB.Close
'''                   AdoAuxDB.MoveNext
'''                Loop
'''             End If
'''             AdoAuxDB.Close
        Case "Trans_Compras"
             Update_Default_SP "Trans_Compras"
        Case "Trans_Kardex"
             Update_Default_SP "Trans_Kardex"
        Case "Trans_SubCtas"
             Update_Default_SP "Trans_SubCtas"
        Case "Trans_Ventas"
             Update_Default_SP "Trans_Ventas"
        Case "Transacciones"
             Update_Default_SP "Transacciones"
      End Select
  Next I
'''  Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
'''  Progreso_Barra.Mensaje_Box = "FIN DEL PROCESO DE ACTUALIZACION"
'''  Progreso_Final
 'MsgBox "Proceso Terminado"
End Sub

Public Sub UPD_Listar_Tablas(Update_LstTablas As ListBox)
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim IJ As Long
Dim IdTime As Long
Dim strCnn As String
Dim NúmeroError
On Error GoTo Errorhandler
' Consultamos las cuentas de la tabla
  RatonReloj
 'MsgBox AdoStrCnn
  Update_LstTablas.Clear
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.open AdoStrCnn
  Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
  Do Until RstSchema.EOF
     If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
        Update_LstTablas.AddItem RstSchema!TABLE_NAME
     End If
     RstSchema.MoveNext
  Loop
  AdoCon1.Close
  RatonNormal
  Exit Sub
Errorhandler:
    RatonNormal
    MsgBox "Error:(" & Err & ")" & vbCrLf _
         & "En la Impresora: " & Printer.DeviceName & vbCrLf _
         & "No pudo imprimir correctamente"
    Exit Sub
End Sub

'''Public Sub Actualizar_Clientes_DBF(Campor_DBF As String, _
'''                                   Base_Datos_DBF As String, _
'''                                   AdoCliente As Adodc, _
'''                                   Tipo_BaseDBF As String)
'''Dim Idx_Espe As Integer
'''Dim Cont_DBF As Integer
'''Dim TDBF_A As Tipo_DBF_Alumnos
'''Dim IdCurso As Integer
'''Dim IdCursoA As Integer
'''Dim Codigo_DBF As String
'''Dim Actualiza_Cliente As Boolean
'''Dim Grupos_Activos As String
'''Dim No_Encontro_Grupo As Boolean
'''
''' If Len(Base_Datos_DBF) > 1 Then
'''   'MsgBox Base_Datos_DBF
'''    Progreso_Barra.Mensaje_Box = "IMPORTANDO DATOS DE " & Tipo_BaseDBF & ": " & Base_Datos_DBF
'''    Progreso_Iniciar
'''
'''    sSQL = "UPDATE Clientes " _
'''         & "SET Grupo = 'RETIRADO' " _
'''         & "WHERE FA <> " & Val(adFalse) & " "
'''    Ejecutar_SQL_SP sSQL
'''    Progreso_Esperar
'''
'''   'Consultamos los campos de la base externa
'''    sSQL = "SELECT " & Campor_DBF & " " _
'''         & "FROM " & Base_Datos_DBF & " " _
'''         & "WHERE cedula <> '.' " _
'''         & "ORDER BY codbanco "
'''    Select_Data_DBF sSQL
'''   'MsgBox Base_Datos_DBF & "..."
'''    With Dato_DBF.Registo
'''     If .RecordCount > 0 Then
'''         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo + (.RecordCount * 2)
'''         Do While Not .EOF
'''            RatonReloj
'''           'Aprobado: S si aprobo el año en caso de antiguos
'''            Progreso_Barra.Mensaje_Box = Tipo_BaseDBF & " DBF: " & Base_Datos_DBF & " => " & .Fields("nombres")
'''            Progreso_Esperar
'''            If IsNull(.Fields("cedular")) Then TDBF_A.cedular = Ninguno Else TDBF_A.cedular = TrimStrg(.Fields("cedular"))
'''            If IsNull(.Fields("pagador")) Then TDBF_A.pagador = Ninguno Else TDBF_A.pagador = TrimStrg(.Fields("pagador"))
'''            If IsNull(.Fields("fonopaga")) Then TDBF_A.fonopaga = Ninguno Else TDBF_A.fonopaga = TrimStrg(.Fields("fonopaga"))
'''            If IsNull(.Fields("direcpaga")) Then TDBF_A.direcpaga = Ninguno Else TDBF_A.direcpaga = TrimStrg(.Fields("direcpaga"))
'''            If IsNull(.Fields("emailpaga")) Then TDBF_A.emailpaga = Ninguno Else TDBF_A.emailpaga = TrimStrg(.Fields("emailpaga"))
'''            If IsNull(.Fields("bus")) Then TDBF_A.bus = Ninguno Else TDBF_A.bus = TrimStrg(.Fields("bus"))
'''            If IsNull(.Fields("aprobado")) Then
'''               TDBF_A.Aprobado = False
'''            Else
'''               If TrimStrg(.Fields("aprobado")) = "S" Then TDBF_A.Aprobado = True Else TDBF_A.Aprobado = False
'''            End If
'''            If IsNull(.Fields("retirado")) Then
'''               TDBF_A.retirado = False
'''            Else
'''               If TrimStrg(.Fields("retirado")) = "S" Then TDBF_A.retirado = True Else TDBF_A.retirado = False
'''            End If
'''
'''            If IsNull(.Fields("pagado")) Then
'''               TDBF_A.pagado = False
'''            Else
'''               If TrimStrg(.Fields("pagado")) = "S" Then TDBF_A.pagado = True Else TDBF_A.pagado = False
'''            End If
'''            If Base_Datos_DBF = Dato_DBF.Actuales Then TDBF_A.pagado = True
'''            If TDBF_A.retirado Then TDBF_A.Estado = "A" Else TDBF_A.Estado = "N"
'''
'''           'TDBF_A.codest = TrimStrg(.Fields("codest"))
'''            TDBF_A.cedula = TrimStrg(.Fields("cedula"))
'''            TDBF_A.Sexo = TrimStrg(.Fields("sexo"))
'''            TDBF_A.Nombres = TrimStrg(.Fields("nombres")) 'solo lectura
'''
'''           'Averiguamos que tipo de Alumno es:
'''            DigVerif = Digito_Verificador(TDBF_A.cedula)
'''            TDBF_A.codest = Tipo_RUC_CI.Codigo_RUC_CI
'''
'''           'Averiguamos que tipo de Beneficiario es:
'''            DigVerif = Digito_Verificador(TDBF_A.cedular)
'''            TDBF_A.TB = Tipo_RUC_CI.Tipo_Beneficiario
'''            Codigo_DBF = TDBF_A.codest
'''
'''           'MsgBox TDBF_A.Nombres
'''           'Grupo del Curso
'''            IdCurso = .Fields("curso")
'''
'''            TDBF_A.Curso = TrimStrg(Format$(IdCurso, "00") & "." & .Fields("especial") & .Fields("paralelo"))  ' Es numerico 1 - 13
'''
'''            TDBF_A.nomcurso = DBF_Cursos(IdCurso)
'''            For Idx_Espe = 1 To UBound(DBF_Especialidad)
'''                If .Fields("especial") = DBF_Especialidad(Idx_Espe).CodEspe Then
'''                    TDBF_A.nomcurso = TDBF_A.nomcurso & " " & DBF_Especialidad(Idx_Espe).Especialidad
'''                    Idx_Espe = UBound(DBF_Especialidad) + 1
'''                End If
'''            Next Idx_Espe
'''
'''           ' No_Encontro_Grupo = True
'''''            For IdCurso = 0 To UBound(DBF_Grupo) - 1
'''''                If TDBF_A.Curso = DBF_Grupo(IdCurso) Then
'''''                  ' No_Encontro_Grupo = False
'''''                   Exit For
'''''                End If
'''''            Next IdCurso
'''''            If No_Encontro_Grupo Then
'''''
'''''               For IdCurso = 0 To UBound(DBF_Grupo) - 1
'''''                   If DBF_Grupo(IdCurso) = Ninguno Then
'''''                      DBF_Grupo(IdCurso) = TDBF_A.Curso
'''''                      IdCurso = UBound(DBF_Grupo)
'''''                   End If
'''''               Next IdCurso
'''''            End If
'''
'''            TDBF_A.nomcurso = TDBF_A.nomcurso & " " & .Fields("paralelo")
'''            TDBF_A.fonopaga = Replace(TDBF_A.fonopaga, "'", "")
'''            TDBF_A.pagador = Replace(TDBF_A.pagador, "'", "")
'''            TDBF_A.direcpaga = Replace(TDBF_A.direcpaga, "'", "")
'''            TDBF_A.cedular = Replace(TDBF_A.cedular, "'", "")
'''            TDBF_A.TB = Replace(TDBF_A.TB, "'", "")
'''            TDBF_A.Curso = Replace(TDBF_A.Curso, "'", "")
'''
'''            TDBF_A.fonopaga = Replace(TDBF_A.fonopaga, "#", "")
'''            TDBF_A.pagador = Replace(TDBF_A.pagador, "#", "")
'''            TDBF_A.direcpaga = Replace(TDBF_A.direcpaga, "#", "")
'''            TDBF_A.cedular = Replace(TDBF_A.cedular, "#", "")
'''            TDBF_A.TB = Replace(TDBF_A.TB, "#", "")
'''            TDBF_A.Curso = Replace(TDBF_A.Curso, "#", "")
'''
'''            If TDBF_A.fonopaga = "" Then TDBF_A.fonopaga = Ninguno
'''            If TDBF_A.pagador = "" Then TDBF_A.pagador = Ninguno
'''            If TDBF_A.direcpaga = "" Then TDBF_A.direcpaga = Ninguno
'''            If TDBF_A.cedular = "" Then TDBF_A.cedular = Ninguno
'''            If TDBF_A.TB = "" Then TDBF_A.TB = Ninguno
'''            If TDBF_A.Curso = "" Then TDBF_A.Curso = Ninguno
'''
'''            If Dato_DBF.Actuales = Base_Datos_DBF Then TDBF_A.pagado = True
'''
'''           'MsgBox TDBF_A.codbanco & vbCrLf & TDBF_A.Nombres & vbCrLf & TDBF_A.Curso & vbCrLf & TDBF_A.nomcurso
'''           'Clientes
'''            Progreso_Barra.Mensaje_Box = Tipo_BaseDBF & " DBF: " & Base_Datos_DBF & " => " & TDBF_A.Nombres
'''            Progreso_Esperar
'''            Actualiza_Cliente = False
'''            sSQL = "SELECT T,FA,Cliente,CI_RUC,TD,Direccion,Grupo,CI_RUC_SRI,TD_SRI,Cedula,Telefono,Email," _
'''                 & "Representante,DireccionT,Casilla,Codigo,Sexo " _
'''                 & "FROM Clientes " _
'''                 & "WHERE CI_RUC = '" & TDBF_A.cedula & "' "
'''            Select_Adodc AdoCliente, sSQL
'''            If AdoCliente.Recordset.RecordCount > 0 Then
'''
'''               'If TDBF_A.Nombres = "FUENMAYOR JAYA BIANCA ROMINA" Then MsgBox TDBF_A.Nombres & ", Retirado: " & TDBF_A.retirado
'''
'''               TDBF_A.codest = AdoCliente.Recordset.Fields("Codigo")
'''              'MsgBox TDBF_A.Estado
'''               If AdoCliente.Recordset.Fields("T") <> TDBF_A.Estado Then
'''                  AdoCliente.Recordset.Fields("T") = TDBF_A.Estado
'''                  Actualiza_Cliente = True
'''               End If
'''               If AdoCliente.Recordset.Fields("Cliente") <> TDBF_A.Nombres Then
'''                  AdoCliente.Recordset.Fields("Cliente") = TDBF_A.Nombres
'''                  Actualiza_Cliente = True
'''               End If
'''               If AdoCliente.Recordset.Fields("Direccion") <> TDBF_A.nomcurso Then
'''                  AdoCliente.Recordset.Fields("Direccion") = UCaseStrg(TDBF_A.nomcurso)
'''                  Actualiza_Cliente = True
'''               End If
'''               If AdoCliente.Recordset.Fields("Grupo") <> TDBF_A.Curso Then
'''                  AdoCliente.Recordset.Fields("Grupo") = TDBF_A.Curso
'''                  Actualiza_Cliente = True
'''               End If
'''               If AdoCliente.Recordset.Fields("Sexo") <> TDBF_A.Sexo Then
'''                  AdoCliente.Recordset.Fields("Sexo") = TDBF_A.Sexo
'''                  Actualiza_Cliente = True
'''               End If
'''               If Len(TDBF_A.cedular) > 1 And AdoCliente.Recordset.Fields("CI_RUC_SRI") <> TDBF_A.cedular Then
'''                  AdoCliente.Recordset.Fields("Cedula") = TDBF_A.cedular
'''                  AdoCliente.Recordset.Fields("CI_RUC_SRI") = TDBF_A.cedular
'''                  AdoCliente.Recordset.Fields("TD_SRI") = TDBF_A.TB
'''                  Actualiza_Cliente = True
'''               End If
'''               If Len(TDBF_A.fonopaga) > 1 And AdoCliente.Recordset.Fields("Telefono") <> TDBF_A.fonopaga Then
'''                  AdoCliente.Recordset.Fields("Telefono") = TDBF_A.fonopaga
'''                  Actualiza_Cliente = True
'''               End If
'''               If Len(TDBF_A.pagador) > 1 And AdoCliente.Recordset.Fields("Representante") <> TDBF_A.pagador Then
'''                  AdoCliente.Recordset.Fields("Representante") = TDBF_A.pagador
'''                  Actualiza_Cliente = True
'''               End If
'''               If Len(TDBF_A.direcpaga) > 1 And AdoCliente.Recordset.Fields("DireccionT") <> TDBF_A.direcpaga Then
'''                  AdoCliente.Recordset.Fields("DireccionT") = TDBF_A.direcpaga
'''                  Actualiza_Cliente = True
'''               End If
'''               If Len(TDBF_A.emailpaga) > 1 And AdoCliente.Recordset.Fields("Email") <> LCase(TDBF_A.emailpaga) Then
'''                  AdoCliente.Recordset.Fields("Email") = LCase(TDBF_A.emailpaga)
'''                  Actualiza_Cliente = True
'''               End If
'''               If TDBF_A.retirado Then
'''                  AdoCliente.Recordset.Fields("Casilla") = "RETIRADO"
'''                  AdoCliente.Recordset.Fields("Grupo") = TDBF_A.Curso
'''                  TextoImprimio = TextoImprimio & "CURSO: " & TDBF_A.Curso & ", NOMBRE: " & TDBF_A.Nombres & vbCrLf
'''                  Actualiza_Cliente = True
'''               End If
''''''               If TDBF_A.retirado Then
''''''                  AdoCliente.Recordset.Fields("Grupo") = "RETIRADO"
''''''                  Actualiza_Cliente = True
''''''               End If
'''
'''               If AdoCliente.Recordset.Fields("FA") <> TDBF_A.pagado Then
'''                  AdoCliente.Recordset.Fields("FA") = True  'TDBF_A.pagado
'''                  Actualiza_Cliente = True
'''               End If
'''
'''               If TDBF_A.Aprobado Then
'''                  If Tipo_BaseDBF = "Antiguos" Then
'''                     For IdCurso = 0 To UBound(DBF_Grupo) - 1
'''                         If TDBF_A.Curso = DBF_Grupo(IdCurso) Then
'''                            IdCursoA = IdCurso
'''                            Exit For
'''                         End If
'''                     Next IdCurso
'''                     Cadena = DBF_Grupo(IdCurso)
'''                     For IdCurso = IdCursoA To UBound(DBF_Grupo) - 1
'''                         If MidStrg(Cadena, 1, 2) <> MidStrg(DBF_Grupo(IdCurso), 1, 2) Then
'''                            TDBF_A.Curso = TrimStrg(MidStrg(DBF_Grupo(IdCurso), 1, 2) & MidStrg(Cadena, 3, Len(Cadena)))
'''                            Exit For
'''                         End If
'''                     Next IdCurso
'''                     IdCurso = Val(MidStrg(TDBF_A.Curso, 1, 2))
'''                     TDBF_A.nomcurso = DBF_Cursos(IdCurso)
'''                     For Idx_Espe = 1 To UBound(DBF_Especialidad)
'''                         If .Fields("especial") = DBF_Especialidad(Idx_Espe).CodEspe Then
'''                             TDBF_A.nomcurso = TDBF_A.nomcurso & " " & DBF_Especialidad(Idx_Espe).Especialidad
'''                             Idx_Espe = UBound(DBF_Especialidad) + 1
'''                         End If
'''                     Next Idx_Espe
'''                     TDBF_A.nomcurso = TDBF_A.nomcurso & " " & .Fields("paralelo")
'''                  End If
'''                  AdoCliente.Recordset.Fields("Grupo") = TDBF_A.Curso
'''                  Actualiza_Cliente = True
'''               End If
'''              'MsgBox Actualiza_Cliente
'''               If Actualiza_Cliente Then AdoCliente.Recordset.Update
'''            Else
'''               Actualiza_Cliente = False
'''               sSQL = "SELECT T,FA,Cliente,CI_RUC,TD,Direccion,Grupo,CI_RUC_SRI,TD_SRI,Cedula,Telefono,Email," _
'''                    & "Representante,DireccionT,Casilla,Codigo,Sexo " _
'''                    & "FROM Clientes " _
'''                    & "WHERE Codigo = '" & TDBF_A.codest & "' "
'''               Select_Adodc AdoCliente, sSQL
'''               If AdoCliente.Recordset.RecordCount > 0 Then
'''                  If AdoCliente.Recordset.Fields("T") <> TDBF_A.Estado Then
'''                     AdoCliente.Recordset.Fields("T") = TDBF_A.Estado
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If AdoCliente.Recordset.Fields("Direccion") <> TDBF_A.nomcurso Then
'''                     AdoCliente.Recordset.Fields("Direccion") = UCaseStrg(TDBF_A.nomcurso)
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If AdoCliente.Recordset.Fields("Grupo") <> TDBF_A.Curso Then
'''                     AdoCliente.Recordset.Fields("Grupo") = TDBF_A.Curso
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If AdoCliente.Recordset.Fields("Sexo") <> TDBF_A.Sexo Then
'''                     AdoCliente.Recordset.Fields("Sexo") = TDBF_A.Sexo
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If Len(TDBF_A.cedular) > 1 And AdoCliente.Recordset.Fields("CI_RUC_SRI") <> TDBF_A.cedular Then
'''                     AdoCliente.Recordset.Fields("Cedula") = TDBF_A.cedular
'''                     AdoCliente.Recordset.Fields("CI_RUC_SRI") = TDBF_A.cedular
'''                     AdoCliente.Recordset.Fields("TD_SRI") = TDBF_A.TB
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If Len(TDBF_A.fonopaga) > 1 And AdoCliente.Recordset.Fields("Telefono") <> TDBF_A.fonopaga Then
'''                     AdoCliente.Recordset.Fields("Telefono") = TDBF_A.fonopaga
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If Len(TDBF_A.pagador) > 1 And AdoCliente.Recordset.Fields("Representante") <> TDBF_A.pagador Then
'''                     AdoCliente.Recordset.Fields("Representante") = TDBF_A.pagador
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If Len(TDBF_A.direcpaga) > 1 And AdoCliente.Recordset.Fields("DireccionT") <> TDBF_A.direcpaga Then
'''                     AdoCliente.Recordset.Fields("DireccionT") = TDBF_A.direcpaga
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If Len(TDBF_A.emailpaga) > 1 And AdoCliente.Recordset.Fields("Email") <> LCase(TDBF_A.emailpaga) Then
'''                     AdoCliente.Recordset.Fields("Email") = LCase(TDBF_A.emailpaga)
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If TDBF_A.retirado Then
'''                     AdoCliente.Recordset.Fields("Casilla") = "RETIRADO"
'''                     AdoCliente.Recordset.Fields("Grupo") = TDBF_A.Curso
'''                     TextoImprimio = TextoImprimio & "CURSO: " & TDBF_A.Curso & ", NOMBRE: " & TDBF_A.Nombres & vbCrLf
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If AdoCliente.Recordset.Fields("FA") <> TDBF_A.pagado Then
'''                     AdoCliente.Recordset.Fields("FA") = True ' TDBF_A.pagado
'''                     Actualiza_Cliente = True
'''                  End If
'''                  If AdoCliente.Recordset.Fields("Cliente") <> TDBF_A.Nombres Then
'''                     TDBF_A.codest = Tipo_RUC_CI.Codigo_RUC_CI
'''                     GoTo Crear_Nuevo
'''                  End If
'''                 'If TDBF_A.Nombres = "BUENAÑO MONTALVO BRITHANY LIZBETH" Then MsgBox "Pare"
'''                  If Actualiza_Cliente Then AdoCliente.Recordset.Update
'''               Else
'''Crear_Nuevo:
'''
'''                   sSQL = "SELECT T,FA,Cliente,CI_RUC,TD,Direccion,Grupo,CI_RUC_SRI,TD_SRI,Cedula,Telefono,Email," _
'''                        & "Representante,DireccionT,Casilla,Codigo,Sexo " _
'''                        & "FROM Clientes " _
'''                        & "WHERE Cliente = '" & TDBF_A.Nombres & "' "
'''                   Select_Adodc AdoCliente, sSQL
'''                   If AdoCliente.Recordset.RecordCount > 0 Then
'''                      TDBF_A.codest = AdoCliente.Recordset.Fields("Codigo")
'''                      If AdoCliente.Recordset.Fields("T") <> TDBF_A.Estado Then
'''                         AdoCliente.Recordset.Fields("T") = TDBF_A.Estado
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If AdoCliente.Recordset.Fields("Direccion") <> TDBF_A.nomcurso Then
'''                         AdoCliente.Recordset.Fields("Direccion") = UCaseStrg(TDBF_A.nomcurso)
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If AdoCliente.Recordset.Fields("Grupo") <> TDBF_A.Curso Then
'''                         AdoCliente.Recordset.Fields("Grupo") = TDBF_A.Curso
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If AdoCliente.Recordset.Fields("Sexo") <> TDBF_A.Sexo Then
'''                         AdoCliente.Recordset.Fields("Sexo") = TDBF_A.Sexo
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If Len(TDBF_A.cedular) > 1 And AdoCliente.Recordset.Fields("CI_RUC_SRI") <> TDBF_A.cedular Then
'''                         AdoCliente.Recordset.Fields("Cedula") = TDBF_A.cedular
'''                         AdoCliente.Recordset.Fields("CI_RUC_SRI") = TDBF_A.cedular
'''                         AdoCliente.Recordset.Fields("TD_SRI") = TDBF_A.TB
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If Len(TDBF_A.fonopaga) > 1 And AdoCliente.Recordset.Fields("Telefono") <> TDBF_A.fonopaga Then
'''                         AdoCliente.Recordset.Fields("Telefono") = TDBF_A.fonopaga
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If Len(TDBF_A.pagador) > 1 And AdoCliente.Recordset.Fields("Representante") <> TDBF_A.pagador Then
'''                         AdoCliente.Recordset.Fields("Representante") = TDBF_A.pagador
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If Len(TDBF_A.direcpaga) > 1 And AdoCliente.Recordset.Fields("DireccionT") <> TDBF_A.direcpaga Then
'''                         AdoCliente.Recordset.Fields("DireccionT") = TDBF_A.direcpaga
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If Len(TDBF_A.emailpaga) > 1 And AdoCliente.Recordset.Fields("Email") <> LCase(TDBF_A.emailpaga) Then
'''                         AdoCliente.Recordset.Fields("Email") = LCase(TDBF_A.emailpaga)
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If TDBF_A.retirado Then
'''                         AdoCliente.Recordset.Fields("Casilla") = "RETIRADO"
'''                         TextoImprimio = TextoImprimio & "CURSO: " & TDBF_A.Curso & ", NOMBRE: " & TDBF_A.Nombres & vbCrLf
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If AdoCliente.Recordset.Fields("FA") <> TDBF_A.pagado Then
'''                         AdoCliente.Recordset.Fields("FA") = True 'TDBF_A.pagado
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If TDBF_A.retirado Then
'''                         AdoCliente.Recordset.Fields("Grupo") = "RETIRADO"
'''                         Actualiza_Cliente = True
'''                      End If
'''                      If Actualiza_Cliente Then AdoCliente.Recordset.Update
'''                   Else
'''                      SetAdoAddNew "Clientes"
'''                      SetAdoFields "T", Normal
'''                      SetAdoFields "TD", Normal
'''                      SetAdoFields "Codigo", TDBF_A.codest
'''                      SetAdoFields "CI_RUC", TDBF_A.cedula
'''                      SetAdoFields "FA", True   '  TDBF_A.retirado
'''                      SetAdoFields "Cliente", TDBF_A.Nombres
'''                      SetAdoFields "Direccion", UCaseStrg(TDBF_A.nomcurso)
'''                      SetAdoFields "Grupo", TDBF_A.Curso
'''                      SetAdoFields "Sexo", TDBF_A.Sexo
'''                      If Len(TDBF_A.cedular) > 1 Then
'''                         SetAdoFields "Cedula", TDBF_A.cedular
'''                         SetAdoFields "CI_RUC_SRI", TDBF_A.cedular
'''                         SetAdoFields "TD_SRI", TDBF_A.TB
'''                      End If
'''                      If Len(TDBF_A.fonopaga) > 1 Then SetAdoFields "Telefono", TDBF_A.fonopaga
'''                      If Len(TDBF_A.pagador) > 1 Then SetAdoFields "Representante", TDBF_A.pagador
'''                      If Len(TDBF_A.direcpaga) > 1 Then SetAdoFields "DireccionT", TDBF_A.direcpaga
'''                      SetAdoUpdate
'''                   End If
'''               End If
'''            End If
'''
'''           'Clientes_Matriculas
'''            Actualiza_Cliente = False
'''            sSQL = "SELECT T,Telefono_R,Representante,Lugar_Trabajo_R,Cedula_R,TD,Grupo_No,Item,Periodo,Codigo " _
'''                 & "FROM Clientes_Matriculas " _
'''                 & "WHERE Periodo = '" & Periodo_Contable & "' " _
'''                 & "AND Item = '" & NumEmpresa & "' " _
'''                 & "AND Codigo = '" & TDBF_A.codest & "' "
'''            Select_Adodc AdoCliente, sSQL
'''            If AdoCliente.Recordset.RecordCount > 0 Then
'''                If AdoCliente.Recordset.Fields("T") <> TDBF_A.Estado Then
'''                   AdoCliente.Recordset.Fields("T") = TDBF_A.Estado
'''                   Actualiza_Cliente = True
'''                End If
'''                If Len(TDBF_A.fonopaga) > 1 And AdoCliente.Recordset.Fields("Telefono_R") <> TDBF_A.fonopaga Then
'''                   AdoCliente.Recordset.Fields("Telefono_R") = TDBF_A.fonopaga
'''                   Actualiza_Cliente = True
'''                End If
'''                If Len(TDBF_A.pagador) > 1 And AdoCliente.Recordset.Fields("Representante") <> TDBF_A.pagador Then
'''                   AdoCliente.Recordset.Fields("Representante") = TDBF_A.pagador
'''                   Actualiza_Cliente = True
'''                End If
'''                If Len(TDBF_A.direcpaga) > 1 And AdoCliente.Recordset.Fields("Lugar_Trabajo_R") <> TDBF_A.direcpaga Then
'''                   AdoCliente.Recordset.Fields("Lugar_Trabajo_R") = TDBF_A.direcpaga
'''                   Actualiza_Cliente = True
'''                End If
'''                If AdoCliente.Recordset.Fields("Cedula_R") <> TDBF_A.cedular Then
'''                   AdoCliente.Recordset.Fields("Cedula_R") = TDBF_A.cedular
'''                   AdoCliente.Recordset.Fields("TD") = TDBF_A.TB
'''                   Actualiza_Cliente = True
'''                End If
'''                If AdoCliente.Recordset.Fields("Grupo_No") <> TDBF_A.Curso Then
'''                   AdoCliente.Recordset.Fields("Grupo_No") = TDBF_A.Curso
'''                   Actualiza_Cliente = True
'''                End If
'''                If Actualiza_Cliente Then AdoCliente.Recordset.Update
'''            Else
'''                SetAdoAddNew "Clientes_Matriculas"
'''                SetAdoFields "T", TDBF_A.Estado
'''                SetAdoFields "Codigo", TDBF_A.codest
'''                SetAdoFields "TD", TDBF_A.TB
'''                SetAdoFields "Grupo_No", TDBF_A.Curso
'''                If Len(TDBF_A.fonopaga) > 1 Then SetAdoFields "Telefono_R", TDBF_A.fonopaga
'''                If Len(TDBF_A.pagador) > 1 Then SetAdoFields "Representante", TDBF_A.pagador
'''                If Len(TDBF_A.direcpaga) > 1 Then SetAdoFields "Lugar_Trabajo_R", TDBF_A.direcpaga
'''                If Len(TDBF_A.cedular) > 1 Then SetAdoFields "Cedula_R", TDBF_A.cedular
'''                SetAdoFields "Periodo", Periodo_Contable
'''                SetAdoFields "Item", NumEmpresa
'''                SetAdoUpdate
'''            End If
'''            Cont_DBF = Cont_DBF + 1
'''            RatonNormal
'''           .MoveNext
'''         Loop
'''     End If
'''    End With
'''    Close_DBF
'''    Grupos_Activos = ""
'''    For IdCurso = 0 To UBound(DBF_Grupo) - 1
'''        If DBF_Grupo(IdCurso) <> Ninguno Then Grupos_Activos = Grupos_Activos & "'" & DBF_Grupo(IdCurso) & "', "
'''    Next IdCurso
'''    Grupos_Activos = Grupos_Activos & "'@'"
'''
'''    sSQL = "UPDATE Clientes " _
'''         & "SET Grupo = 'RETIRADO' " _
'''         & "WHERE Grupo NOT IN (" & Grupos_Activos & ") "
'''    Ejecutar_SQL_SP sSQL
'''    Progreso_Final
''' End If
'''End Sub

Public Function Full_Fields(NombreTabla As String) As String
Dim AdoTableDB As ADODB.Recordset
Dim SQLTable As String
Dim ListaCampos As String

    ListaCampos = ""
    If Len(NombreTabla) > 1 Then
       SQLTable = "SELECT COL.name " _
                & "FROM dbo.syscolumns COL " _
                & "JOIN dbo.sysobjects OBJ ON OBJ.id = COL.id " _
                & "WHERE OBJ.name = '" & NombreTabla & "' " _
                & "AND (OBJ.xtype='U' OR OBJ.xtype='V') " _
                & "ORDER BY COL.colid "
       Select_AdoDB AdoTableDB, SQLTable
       With AdoTableDB
        If .RecordCount > 0 Then
            Do While Not .EOF
               ListaCampos = ListaCampos & .fields("name") & ", "
              .MoveNext
            Loop
            ListaCampos = TrimStrg(ListaCampos)
            ListaCampos = MidStrg(ListaCampos, 1, Len(ListaCampos) - 1)
        End If
       End With
       AdoTableDB.Close
    End If
    Full_Fields = ListaCampos
End Function

Public Function Leer_Datos_Clientes(Codigo_CIRUC_Cliente As String, Optional NoActualizaSP As Boolean) As Tipo_Beneficiarios
Dim AdoCliDB As ADODB.Recordset
Dim TBenef As Tipo_Beneficiarios
Dim Por_Codigo As Boolean
Dim Por_CIRUC As Boolean
Dim Por_Cliente As Boolean
Dim CadAux As String

    Por_Codigo = False
    Por_CIRUC = False
    Por_Cliente = False
    With TBenef
        .FA = False
        .Asignar_Dr = False
        .Codigo = Ninguno
        .Cliente = Ninguno
        .Tipo_Cta = Ninguno
        .Cta_Numero = Ninguno
        .Descuento = False
        .T = ""
        .TP = ""
        .CI_RUC = ""
        .TD = ""
        .Fecha = ""
        .Fecha_A = ""
        .Fecha_N = ""
        .Sexo = ""
        .Email1 = ""
        .Email2 = ""
        .Direccion = ""
        .DirNumero = ""
        .Telefono = ""
        .Telefono1 = ""
        .TelefonoT = ""
        .Celular = ""
        .Ciudad = ""
        .Prov = ""
        .Pais = ""
        .Profesion = ""
        .Archivo_Foto = ""
        .Representante = Ninguno
        .RUC_CI_Rep = ""
        .TD_Rep = ""
        .Direccion_Rep = "SD"
        .Grupo_No = ""
        .Contacto = ""
        .Calificacion = ""
        .Plan_Afiliado = ""
        .Cte_Ahr_Otro = ""
        .Cta_Transf = ""
        .Cod_Banco = 0
        .Salario = 0
        .Saldo_Pendiente = 0
        .Total_Anticipo = 0
    End With
    
    If Len(Codigo_CIRUC_Cliente) <= 0 Then Codigo_CIRUC_Cliente = Ninguno
    'MsgBox Codigo_CIRUC_Cliente
    If Not NoActualizaSP Then Leer_Datos_Cliente_SP Codigo_CIRUC_Cliente
    'MsgBox Codigo_CIRUC_Cliente
    TBenef.Codigo = Codigo_CIRUC_Cliente
        
   'Verificamos la informacion del Clienete
    If TBenef.Codigo <> "." Then
       With TBenef
            sSQL = "SELECT " & Full_Fields("Clientes") & " " _
                 & "FROM Clientes " _
                 & "WHERE Codigo = '" & .Codigo & "' "
            Select_AdoDB AdoCliDB, sSQL
            If AdoCliDB.RecordCount > 0 Then
              .FA = AdoCliDB.fields("FA")
              .Asignar_Dr = AdoCliDB.fields("Asignar_Dr")
              .Cliente = AdoCliDB.fields("Cliente")
              .Descuento = AdoCliDB.fields("Descuento")
              .T = AdoCliDB.fields("T")
              .CI_RUC = AdoCliDB.fields("CI_RUC")
              .TD = AdoCliDB.fields("TD")
              .Fecha = AdoCliDB.fields("Fecha")
              .Fecha_N = AdoCliDB.fields("Fecha_N")
              .Sexo = AdoCliDB.fields("Sexo")
              .Email1 = AdoCliDB.fields("Email")
              .Email2 = AdoCliDB.fields("Email2")
              .EmailR = AdoCliDB.fields("EmailR")
              .Direccion = AdoCliDB.fields("Direccion")
              .DirNumero = AdoCliDB.fields("DirNumero")
              .Telefono = AdoCliDB.fields("Telefono")
              .Telefono1 = AdoCliDB.fields("Telefono_R")
              .TelefonoT = AdoCliDB.fields("TelefonoT")
              .Ciudad = AdoCliDB.fields("Ciudad")
              .Prov = AdoCliDB.fields("Prov")
              .Pais = AdoCliDB.fields("Pais")
              .Profesion = AdoCliDB.fields("Profesion")
              .Grupo_No = AdoCliDB.fields("Grupo")
              .Contacto = AdoCliDB.fields("Contacto")
              .Calificacion = AdoCliDB.fields("Calificacion")
              .Plan_Afiliado = AdoCliDB.fields("Plan_Afiliado")
              .Actividad = AdoCliDB.fields("Actividad")
              .Credito = AdoCliDB.fields("Credito")
              
              .Representante = Replace(AdoCliDB.fields("Representante"), "  ", " ")
              .RUC_CI_Rep = AdoCliDB.fields("CI_RUC_R")
              .TD_Rep = AdoCliDB.fields("TD_R")
              .Tipo_Cta = AdoCliDB.fields("Tipo_Cta")
              .Cod_Banco = AdoCliDB.fields("Cod_Banco")
              .Cta_Numero = AdoCliDB.fields("Cta_Numero")
              .Direccion_Rep = AdoCliDB.fields("DireccionT")
              .Fecha_Cad = AdoCliDB.fields("Fecha_Cad")
              .Saldo_Pendiente = AdoCliDB.fields("Saldo_Pendiente")
              .Archivo_Foto = AdoCliDB.fields("Archivo_Foto")
             '.Salario = 0
            End If
            AdoCliDB.Close
       End With
    End If
    Leer_Datos_Clientes = TBenef
End Function

Public Sub Conectar_Base_Datos()
Dim CarBase As String
Dim Conexion_Temp As String
Dim Nombre_Base_Datos_SQL As String
Dim CarIni As Long
Dim CarFin As Long

  'Leemos la cadena de conexion
   RutaGeneraFile = RutaSistema & "\ConectarDB.ini"
   RutaEmpresa = RutaSistema & "\EMPRESA\"
   Conexion_Temp = ""
   Nombre_Base_Datos_SQL = ""
   strNombreBaseDatos = ""
   strWebServices = "000"
   strPuerto = "1433"
   strIPServidor = "127.0.0.1"
  'Leemos archivo de conexion
   NumFile = FreeFile
   Open RutaGeneraFile For Input As #NumFile
   Do While Not EOF(NumFile)
      CarBase = Input(1, #NumFile) ' Obtiene un carácter.
      Conexion_Temp = Conexion_Temp & CarBase
   Loop
   Close #NumFile
   If InStr(1, Conexion_Temp, "SQLServer = SI") Then SQL_Server = True Else SQL_Server = False
   If InStr(1, Conexion_Temp, "SQLServer = SI") Then
      CarIni = InStr(1, Conexion_Temp, "<SQLServer = SI>") + 18
      CarFin = InStr(1, Conexion_Temp, "</SQLServer>") - 2
      Conexion_Temp = MidStrg(Conexion_Temp, CarIni, CarFin - CarIni)
      
      CarIni = InStr(1, Conexion_Temp, "Data Source=")
      strIPServidor = MidStrg(Conexion_Temp, CarIni + 12, Len(Conexion_Temp))
      CarFin = InStr(1, strIPServidor, ";")
      strIPServidor = MidStrg(strIPServidor, 1, CarFin - 1)
      
      CarIni = InStr(1, strIPServidor, ",")
      CarFin = InStr(1, strIPServidor, ";")
      If CarIni > 0 Then strPuerto = MidStrg(strIPServidor, CarIni + 1, Len(strIPServidor) - CarIni)

      CarIni = InStr(1, strIPServidor, ":")
      CarFin = InStr(1, strIPServidor, ",")
      If CarIni > 0 And CarFin > 0 Then strIPServidor = MidStrg(strIPServidor, CarIni + 1, CarFin - CarIni - 1)
      
      CarIni = InStr(1, Conexion_Temp, "UID=")
      strUsuario = MidStrg(Conexion_Temp, CarIni + 4, Len(Conexion_Temp))
      CarFin = InStr(1, strUsuario, ";")
      strUsuario = MidStrg(strUsuario, 1, CarFin - 1)

      CarIni = InStr(1, Conexion_Temp, "PWD=")
      strPassword = MidStrg(Conexion_Temp, CarIni + 4, Len(Conexion_Temp))
      CarFin = InStr(1, strPassword, ";")
      strPassword = MidStrg(strPassword, 1, CarFin - 1)

      CarIni = InStr(1, Conexion_Temp, "Initial Catalog=")
      strNombreBaseDatos = MidStrg(Conexion_Temp, CarIni + 16, Len(Conexion_Temp))
      CarFin = InStr(1, strNombreBaseDatos, ";")
      strNombreBaseDatos = MidStrg(strNombreBaseDatos, 1, CarFin - 1)
      PathEmpresa = ""
      
      CarIni = InStr(1, Conexion_Temp, "WebServices=")
      strWebServices = MidStrg(Conexion_Temp, CarIni + 12, Len(Conexion_Temp))
      CarFin = InStr(1, strWebServices, ";")
      If CarIni > 0 Then strWebServices = MidStrg(strWebServices, 1, CarFin - 1) Else strWebServices = "000"
   ElseIf InStr(1, Conexion_Temp, "MySQL = SI") Then
      CarIni = InStr(1, Conexion_Temp, "<MySQL = SI") + 13
      CarFin = InStr(1, Conexion_Temp, "</MySQL>") - 2
      Conexion_Temp = MidStrg(Conexion_Temp, CarIni, CarFin - CarIni)
      
      CarIni = InStr(1, Conexion_Temp, "DATABASE=")
      Nombre_Base_Datos_SQL = MidStrg(Conexion_Temp, CarIni, Len(Conexion_Temp))
      CarFin = InStr(1, Nombre_Base_Datos_SQL, ";")
      strNombreBaseDatos = MidStrg(Nombre_Base_Datos_SQL, 10, CarFin - 10)
      PathEmpresa = ""
      
      CarIni = InStr(1, Conexion_Temp, "WebServices=")
      strWebServices = MidStrg(Conexion_Temp, CarIni, Len(Conexion_Temp))
      CarFin = InStr(1, strWebServices, ";")
      If CarIni > 0 Then strWebServices = MidStrg(strWebServices, 13, 3) Else strWebServices = "000"
   Else
      CarIni = InStr(1, Conexion_Temp, "<Access = SI>") + 15
      CarFin = InStr(1, Conexion_Temp, "</Access>") - 2
      Conexion_Temp = MidStrg(Conexion_Temp, CarIni, CarFin - CarIni)
      Conexion_Temp = Replace(Conexion_Temp, "Path\", RutaSistema & "\EMPRESA\")
      
      CarIni = InStr(1, Conexion_Temp, "Data Source=")
      Nombre_Base_Datos_SQL = MidStrg(Conexion_Temp, CarIni, Len(Conexion_Temp))
      CarFin = InStr(1, Nombre_Base_Datos_SQL, ";")
      strNombreBaseDatos = MidStrg(Nombre_Base_Datos_SQL, 13, CarFin - 13)
      PathEmpresa = strNombreBaseDatos
      
      CarIni = InStr(1, Conexion_Temp, "WebServices=")
      strWebServices = MidStrg(Conexion_Temp, CarIni, Len(Conexion_Temp))
      CarFin = InStr(1, strWebServices, ";")
      If CarIni > 0 Then strWebServices = MidStrg(strWebServices, 13, 3) Else strWebServices = "000"
   End If
   AdoStrCnn = Conexion_Temp
'''   MsgBox strIPServidor & vbCrLf _
'''        & strNombreBaseDatos & vbCrLf _
'''        & strWebServices & vbCrLf _
'''        & strPassword & vbCrLf _
'''        & strUsuario & vbCrLf _
'''        & strPuerto & vbCrLf _
'''        & "-----------------------------------------" & vbCrLf _
'''        & AdoStrCnn
End Sub

Public Sub Leer_Datos_Conexion(Conexion As Tipo_Conexion)
Dim LineFile       As String
Dim Conexion_Temp  As String
Dim Linea_Conexion As String
Dim Dato_Cnn       As String

Dim CarIni         As Long
Dim CarFin         As Long

Dim Es_Access     As Boolean
Dim Es_SQL_Server As Boolean
Dim Es_My_SQL     As Boolean

   Es_Access = False
   Es_SQL_Server = False
   Es_My_SQL = False
   With Conexion
       .Tipo_Base = ""
       .Entidad = ""
       .IP_Server = ""
       .Base_Datos = ""
       .Usuario = ""
       .Clave = ""
       .Controlador = ""
       .Opcion = 0
       .Puerto = 0
   End With
  'Leemos la cadena de conexion
   Conexion_Temp = ""
   RutaGeneraFile = RutaSistema & "\ConectarDB.ini"
   NumFile = FreeFile
   Open RutaGeneraFile For Input As #NumFile
   Do While Not EOF(NumFile)
      Line Input #NumFile, LineFile
      If InStr(1, LineFile, "<Access = SI>") Then Es_Access = True
      If InStr(1, LineFile, "<SQLServer = SI>") Then Es_SQL_Server = True
      If InStr(1, LineFile, "<MySQL = SI>") Then Es_My_SQL = True
      Conexion_Temp = Conexion_Temp & LineFile & vbCrLf
   Loop
   Close #NumFile
   SQL_Server = Es_SQL_Server
   If Es_Access Then
      Conexion.Tipo_Base = "ACCESS"
      CarIni = InStr(1, Conexion_Temp, "<Access = SI>") + 15
      CarFin = InStr(1, Conexion_Temp, "</Access>") - 2
      Conexion_Temp = MidStrg(Conexion_Temp, CarIni, CarFin - CarIni)
      Conexion_Temp = Replace(Conexion_Temp, "Path\", RutaSistema & "\EMPRESA\")
   End If
   If Es_SQL_Server Then
      Conexion.Tipo_Base = "SQL SERVER"
      CarIni = InStr(1, Conexion_Temp, "<SQLServer = SI>") + 18
      CarFin = InStr(1, Conexion_Temp, "</SQLServer>") - 2
      Conexion_Temp = MidStrg(Conexion_Temp, CarIni, CarFin - CarIni)
   End If
   If Es_My_SQL Then
      Conexion.Tipo_Base = "MY SQL"
      CarIni = InStr(1, Conexion_Temp, "<MySQL = SI>") + 14
      CarFin = InStr(1, Conexion_Temp, "</MySQL>") - 2
      Conexion_Temp = MidStrg(Conexion_Temp, CarIni, CarFin - CarIni)
   End If
   Do While Len(Conexion_Temp) > 0
      CarFin = InStr(1, Conexion_Temp, ";")
      Linea_Conexion = MidStrg(Conexion_Temp, 1, CarFin - 1)
      CarIni = InStr(1, Linea_Conexion, "=")
      Dato_Cnn = MidStrg(Linea_Conexion, CarIni, Len(Linea_Conexion))
      Select Case MidStrg(Linea_Conexion, 1, CarIni)
        Case "Data Source=", "SERVER=": Conexion.IP_Server = Dato_Cnn
        Case "Provider=", "DRIVER=": Conexion.Controlador = Dato_Cnn
        Case "Initial Catalog=", "DATABASE=": Conexion.Base_Datos = Dato_Cnn
        Case "UID=", "USER=": Conexion.Usuario = Dato_Cnn
        Case "PWD=", "PASSWORD=": Conexion.Clave = Dato_Cnn
        Case "PORT=": Conexion.Puerto = Val(Dato_Cnn)
        Case "OPTION=": Conexion.Opcion = Val(Dato_Cnn)
      End Select
     'MsgBox "LN: " & Linea_Conexion
      Conexion_Temp = TrimStrg(MidStrg(Conexion_Temp, Len(Linea_Conexion) + 2, Len(Conexion_Temp)))
   Loop
   Conexion.Entidad = "CONEXION EMPRESA LOCAL"
   If Es_Access Then
      CarIni = InStr(1, Conexion.IP_Server, "\")
      Conexion.Base_Datos = MidStrg(Conexion.IP_Server, CarIni + 1, Len(Conexion.IP_Server))
      Conexion.IP_Server = RutaSistema & "\EMPRESA\"
   End If
  'MsgBox "..."
End Sub

Public Function Buscar_Beneficiario(PatronBusqueda As String, TipoBus As TipoBusqueda) As String
Dim AdoReg As ADODB.Recordset
Dim CodBenef As String
Dim cSQL As String
  ' X_CI_RUC = 1
  ' X_Beneficiario = 2
  ' X_Grupo = 3
    CodBenef = Ninguno
    CodigoCliente = Ninguno
    NombreCliente = Ninguno
    DireccionCli = Ninguno
    'Grupo = Ninguno
    
    cSQL = "SELECT Codigo,Cliente,Direccion " _
         & "FROM Clientes "
    Select Case TipoBus
      Case 1: cSQL = cSQL & "WHERE CI_RUC = '" & PatronBusqueda & "' "
      Case 2: cSQL = cSQL & "WHERE Cliente = '" & PatronBusqueda & "' "
      Case 3: cSQL = cSQL & "WHERE Grupo = '" & PatronBusqueda & "' "
      Case Else: cSQL = cSQL & "WHERE Codigo = '" & Ninguno & "' "
    End Select
    Set AdoReg = New ADODB.Recordset
    AdoReg.CursorType = adOpenStatic   'adOpenDynamic
    AdoReg.CursorLocation = adUseClient
    AdoReg.LockType = adLockOptimistic
    cSQL = CompilarSQL(cSQL)
    AdoReg.open cSQL, AdoStrCnn, , , adCmdText
    If AdoReg.RecordCount > 0 Then
       CodigoCliente = AdoReg.fields("Codigo")
       NombreCliente = AdoReg.fields("Cliente")
       DireccionCli = AdoReg.fields("Direccion")
       'Grupo = AdoReg.Fields("Grupo")
    End If
    AdoReg.Close
    CodBenef = CodigoCliente
    Buscar_Beneficiario = CodBenef
End Function

'URLInet As Inet,
Public Sub Datos_del_Cliente(C1 As Comprobantes)
Dim Strgs As String
Dim CodigoC As String
Dim AdoRegs As ADODB.Recordset
    C1.TipoContribuyente = ""
    C1.RUC_CI = Ninguno
    C1.Beneficiario = Ninguno
    C1.Email = Ninguno
    C1.TD = Ninguno
    C1.Direccion = Ninguno
    C1.Telefono = Ninguno
    C1.Grupo = Ninguno
    C1.AgenteRetencion = Ninguno
    C1.MicroEmpresa = Ninguno
    C1.Estado = Ninguno
    If IsNull(C1.CodigoB) Or C1.CodigoB = "" Then C1.CodigoB = Ninguno
    
    Strgs = "SELECT TD,CI_RUC,Codigo,Cliente,Email,Grupo,RISE,Especial,Direccion,Telefono " _
          & "FROM Clientes " _
          & "WHERE Codigo = '" & C1.CodigoB & "' "
    Select_AdoDB AdoRegs, Strgs
    With AdoRegs
     If .RecordCount > 0 Then
         C1.RUC_CI = .fields("CI_RUC")
         C1.Beneficiario = .fields("Cliente")
         C1.Email = .fields("Email")
         C1.TD = .fields("TD")
         C1.Direccion = .fields("Direccion")
         C1.Telefono = .fields("Telefono")
         C1.Grupo = .fields("Grupo")
         If .fields("RISE") Then C1.TipoContribuyente = C1.TipoContribuyente & " RISE"
         If .fields("Especial") Then C1.TipoContribuyente = C1.TipoContribuyente & " Contribuyente especial"
         'TipoSRI = consulta_RUC_SRI( C1.RUC_CI)
         If Len(C1.RUC_CI) = 13 Then Tipo_Contribuyente_SP_MySQL C1.RUC_CI, TipoSRI.MicroEmpresa, TipoSRI.AgenteRetencion
         Select Case C1.TD
           Case "C": TipoSRI.Estado = "CEDULA"
           Case "P": TipoSRI.Estado = "PASAPORTE"
           Case "R": TipoSRI.Estado = "RUC ACTIVO"
         End Select
         C1.AgenteRetencion = TipoSRI.AgenteRetencion
         C1.MicroEmpresa = TipoSRI.MicroEmpresa
         C1.Estado = TipoSRI.Estado
     End If
    End With
    AdoRegs.Close
End Sub

'''Public Sub Leer_Datos_FoxPro()
'''    If DirectoryFileExists(Dato_DBF.Carpeta) Then
'''     If Len(Dato_DBF.Curso) > 1 Then
'''        Cont_DBF = 1
'''        sSQL = "SELECT * " _
'''             & "FROM " & Dato_DBF.Curso & " " _
'''             & "WHERE Curso > 0 " _
'''             & "ORDER BY Curso "
'''        Select_Data_DBF sSQL
'''        With Dato_DBF.Registo
'''         If .RecordCount > 0 Then
'''             ReDim DBF_Cursos(.RecordCount + 1) As String
'''             Do While Not .EOF
'''                DBF_Cursos(Cont_DBF) = .Fields("nomcomple")
'''                Cont_DBF = Cont_DBF + 1
'''               .MoveNext
'''             Loop
'''         End If
'''        End With
'''        Close_DBF
'''     End If
'''
'''     If Len(Dato_DBF.Especialidad) > 1 Then
'''        Cont_DBF = 1
'''        sSQL = "SELECT * " _
'''             & "FROM " & Dato_DBF.Especialidad & " " _
'''             & "WHERE espec <> '.' " _
'''             & "ORDER BY espec "
'''        Select_Data_DBF sSQL
'''        With Dato_DBF.Registo
'''         If .RecordCount > 0 Then
'''             ReDim DBF_Especialidad(.RecordCount + 1) As Tipo_Espe_DBF
'''             Do While Not .EOF
'''                DBF_Especialidad(Cont_DBF).CodEspe = .Fields("espec")
'''                DBF_Especialidad(Cont_DBF).Especialidad = .Fields("nomlargo")
'''                Cont_DBF = Cont_DBF + 1
'''               .MoveNext
'''             Loop
'''         End If
'''        End With
'''        Close_DBF
'''     End If
'''
'''     If Len(Dato_DBF.Paralelo) > 1 Then
'''        Cont_DBF = 1
'''        sSQL = "SELECT * " _
'''             & "FROM " & Dato_DBF.Paralelo & " " _
'''             & "WHERE paralelo <> '.' " _
'''             & "ORDER BY paralelo "
'''        Select_Data_DBF sSQL
'''        With Dato_DBF.Registo
'''         If .RecordCount > 0 Then
'''             ReDim DBF_Paralelo(.RecordCount + 1) As String
'''             Do While Not .EOF
'''                DBF_Paralelo(Cont_DBF) = .Fields("paralelo")
'''                Cont_DBF = Cont_DBF + 1
'''               .MoveNext
'''             Loop
'''         End If
'''        End With
'''        Close_DBF
'''     End If
'''
'''    'Consultamos los campos de la base externa
'''     Cont_DBF = 0
'''     sSQL = "SELECT curso,especial,paralelo " _
'''          & "FROM " & Dato_DBF.Antiguos & " " _
'''          & "WHERE aprobado = 'S' " _
'''          & "GROUP BY curso,especial,paralelo " _
'''          & "ORDER BY curso,especial,paralelo "
'''     Select_Data_DBF sSQL
'''     With Dato_DBF.Registo
'''      If .RecordCount > 0 Then
'''          ReDim DBF_Grupo(.RecordCount) As String
'''          For IdCurso = 0 To UBound(DBF_Grupo) - 1
'''              DBF_Grupo(IdCurso) = Ninguno
'''          Next IdCurso
'''
'''          Do While Not .EOF
'''             Cadena = Format$(.Fields("curso"), "00") & "." & .Fields("especial") & .Fields("paralelo")
'''             DBF_Grupo(Cont_DBF) = TrimStrg(Cadena)
'''             Cont_DBF = Cont_DBF + 1
'''            .MoveNext
'''          Loop
'''      End If
'''     End With
'''     Close_DBF
'''  End If
'''End Sub

'''Public Sub Leer_Datos_MySQL()
'''Dim AdoDBCli As ADODB.Recordset
'''Dim RC As MYSQL_CONNECTION
'''Dim RG As MYSQL_RS
'''
'''Dim TBe As Tipo_Beneficiarios
'''
'''   'Activamos conexion del MySQL de base de datos externa
'''    strBaseDatos = Dato_DBF.Actuales
'''    strServidor = Dato_DBF.Carpeta
'''    strUsuario = Dato_DBF.Usuario
'''    strPassword = Dato_DBF.Clave
'''    strPuerto = CStr(Dato_DBF.puerto)
'''
'''    Set RC = New MYSQL_CONNECTION
'''    Set RG = New MYSQL_RS
'''
'''    RC.OpenConnection strServidor, strUsuario, strPassword, strBaseDatos, strPuerto
'''    If RC.State = MY_CONN_OPEN Then
'''       sSQL = "SELECT e.codigoest, e.cedula, CONCAT(e.apellido,' ',e.nombre) As nombreEstudiante, r.celular, r.email, " _
'''            & "r.nombrefactura, r.telfactura, r.dirfactura, r.identificacionfactura, r.tipoidentificacion, r.emailfac " _
'''            & "FROM estudiantes As e, representante As r, matriculas As m, cursos As c, periodo As p " _
'''            & "WHERE e.codigoest = m.Estudiante " _
'''            & "AND e.representante = r.codigorep " _
'''            & "AND c.codigocur = m.curso " _
'''            & "AND c.periodo = p.codigoper " _
'''            & "AND p.descripcion LIKE '%" & Dato_DBF.Periodo & "%' " _
'''            & "ORDER BY  e.apellido, e.nombre,r.emailfac "
'''       sSQL = CompilarSQL(sSQL)
'''       RG.OpenRs sSQL, RC, adOpenKeyset, adLockOptimistic
'''       If RG.RecordCount > 0 Then
'''         'MsgBox "Tabla con registros", vbCritical
'''          Do While Not RG.EOF
'''             Codigo = RG.Fields("codigoest")
'''             sSQL = "SELECT * " _
'''                  & "FROM Clientes_Datos_Externos " _
'''                  & "WHERE Item = '" & NumEmpresa & "' " _
'''                  & "AND Periodo = '" & Periodo_Contable & "' " _
'''                  & "AND codigoest = " & Val(Codigo) & " "
'''             Select_AdoDB AdoDBCli, sSQL
'''             If AdoDBCli.RecordCount = 0 Then
'''                SetAdoAddNew "Clientes_Datos_Externos"
'''                SetAdoFields "Fecha_Act", FechaSistema  'RG.Fields("Fecha_Act")
'''                SetAdoFields "codigoest", RG.Fields("codigoest")
'''                SetAdoFields "cedula", RG.Fields("cedula")
'''                SetAdoFields "nombreEstudiante", RG.Fields("nombreEstudiante")
'''                SetAdoFields "celular", RG.Fields("celular")
'''                SetAdoFields "email", RG.Fields("email")
'''                SetAdoFields "nombrefactura", RG.Fields("nombrefactura")
'''                SetAdoFields "telfactura", RG.Fields("telfactura")
'''                SetAdoFields "dirfactura", RG.Fields("dirfactura")
'''                SetAdoFields "identificacionfactura", RG.Fields("identificacionfactura")
'''                SetAdoFields "tipoidentificacion", RG.Fields("tipoidentificacion")
'''                SetAdoFields "emailfac", RG.Fields("emailfac")
'''
'''                DigVerif = Digito_Verificador(RG.Fields("cedula"))
'''                SetAdoFields "TE", Tipo_RUC_CI.Tipo_Beneficiario
'''                DigVerif = Digito_Verificador(RG.Fields("identificacionfactura"))
'''                SetAdoFields "TR", Tipo_RUC_CI.Tipo_Beneficiario
'''
'''                TBe = Leer_Datos_Clientes(RG.Fields("nombreEstudiante"))
'''                If Len(TBe.Representante) > 1 Then
'''                   SetAdoFields "Representante", TBe.Representante
'''                   SetAdoFields "CI_RUC", TBe.RUC_CI_Rep
'''                   SetAdoFields "TB", TBe.TD_Rep
'''                   SetAdoFields "Email_R", TBe.EmailR
'''                   SetAdoFields "Grupo", TBe.Grupo_No
'''                Else
'''                   SetAdoFields "Representante", "CONSUMIDOR FINAL"
'''                   SetAdoFields "CI_RUC", "9999999999999"
'''                   SetAdoFields "TB", "R"
'''                   SetAdoFields "Email_R", EmailProcesos
'''                End If
'''                SetAdoFields "Periodo", Periodo_Contable
'''                SetAdoFields "Item", NumEmpresa
'''                SetAdoUpdate
'''             End If
'''             AdoDBCli.Close
'''             RG.MoveNext
'''          Loop
'''       Else
'''          MsgBox "tabla en cero", vbCritical
'''       End If
'''       RG.CloseRecordset
'''       RC.CloseConnection
'''    End If
'''
'''   'Restauramos conexion del MySQL de DiskCover
'''    strBaseDatos = "DiskCover_Empresas"
'''    strServidor = "db.diskcoversystem.com"
'''    strUsuario = "diskcover"
'''    strPassword = "disk2017Cover"
'''    strPuerto = "13306"
'''End Sub

Public Sub Optimizar_Memoria()
Dim AdoConexion As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim AdoTabla As ADODB.Recordset

Dim NombreTabla As String
Dim Campos As String
Dim InsCampos As String
Dim LineaTexto As String
Dim Idx As Integer

Dim NumFile As Long

Dim Existe_ID As Boolean

    Progreso_Barra.Incremento = 0
    Progreso_Barra.Valor_Maximo = 100
    Progreso_Barra.Mensaje_Box = "CREANDO SCRIP DE MEMORIA OPTIMIZADA"
    Progreso_Iniciar
    Progreso_Esperar

    RatonReloj
    RutaGeneraFile = RutaSysBases & "\TEMP\SCRIPT_MEMORIA_OPTIMIZADA.sql"
   'FSeteos.Caption = RutaGeneraFile
    LineaTexto = ""
    NumFile = FreeFile
    Open RutaGeneraFile For Output As #NumFile

   'Crea variables de objeto para los objetos de acceso a datos.
    Set AdoConexion = New ADODB.Connection
    AdoConexion.open AdoStrCnn
    Set RstSchema = AdoConexion.OpenSchema(adSchemaTables)
    Do Until RstSchema.EOF
       If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
         'Llenamos la lista de Tablas
          Existe_ID = False
          NombreTabla = RstSchema!TABLE_NAME
          sSQL = "SELECT TOP 1 * " _
               & "FROM " & NombreTabla & " "
          Select_AdoDB AdoTabla, sSQL
          With AdoTabla
               Campos = ""
               For Idx = 0 To .fields.Count - 1
                   If .fields(Idx).Name <> "ID" Then
                       Campos = Campos & "[" & .fields(Idx).Name & "] "
                       If SQL_Server Then
                          Campos = Campos & FieldTypeSQL(.fields(Idx).Type) & " "
                       Else
                          Campos = Campos & FieldTypeAccess(.fields(Idx).Type) & " "
                       End If
                       If .fields(Idx).Type = adVarWChar Then Campos = Campos & "(" & CStr(.fields(Idx).DefinedSize) & ")"
                       If .fields(Idx).Type = adNumeric Then Campos = Campos & "(18,0)"
                       Campos = Campos & " NULL"
                   End If
                   If .fields(Idx).Name = "ID" Then
                       If SQL_Server Then
                          Campos = Campos & "[ID] INT IDENTITY NOT NULL PRIMARY KEY NONCLUSTERED "
                       Else
                          Campos = Campos & "[ID] LONG IDENTITY NOT NULL PRIMARY KEY NONCLUSTERED "
                       End If
                       Existe_ID = True
                   End If
                   Campos = Campos & "," & vbCrLf
               Next Idx
               Campos = MidStrg(Campos, 1, Len(Campos) - 3)
          End With
          AdoTabla.Close
          
          If Existe_ID Then
             sSQL = "CREATE TABLE [T_" & NombreTabla & "] " & vbCrLf _
                  & "(" & Campos & ") " & vbCrLf _
                  & "WITH (MEMORY_OPTIMIZED=ON, DURABILITY=SCHEMA_ONLY);" & vbCrLf _
                  & "GO " & vbCrLf
            'MsgBox "Creando Tabla Nueva:" & vbCrLf & sSQL
             LineaTexto = LineaTexto & sSQL
             Print #NumFile, LineaTexto
          End If
       End If
       RstSchema.MoveNext
    Loop
    RstSchema.Close
    Close #NumFile
    
    RutaGeneraFile = RutaSysBases & "\TEMP\SCRIPT_INSERTAR_DATOS.sql"
   'FSeteos.Caption = RutaGeneraFile
    LineaTexto = ""
    NumFile = FreeFile
    Open RutaGeneraFile For Output As #NumFile
    
    Set RstSchema = AdoConexion.OpenSchema(adSchemaTables)
    Do Until RstSchema.EOF
       If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
         'Llenamos la lista de Tablas
          Existe_ID = False
          NombreTabla = RstSchema!TABLE_NAME
          sSQL = "SELECT TOP 1 * " _
               & "FROM " & NombreTabla & " "
          Select_AdoDB AdoTabla, sSQL
          With AdoTabla
               InsCampos = ""
               For Idx = 0 To .fields.Count - 1
                   If .fields(Idx).Name <> "ID" Then InsCampos = InsCampos & .fields(Idx).Name & ","
                   If .fields(Idx).Name = "ID" Then Existe_ID = True
               Next Idx
          End With
          AdoTabla.Close
          InsCampos = MidStrg(InsCampos, 1, Len(InsCampos) - 1)
          sSQL = "INSERT INTO [T_" & NombreTabla & "] (" & InsCampos & ") " & vbCrLf _
               & "SELEC " & InsCampos & " " _
               & "FROM [" & NombreTabla & "];" & vbCrLf _
               & "GO " & vbCrLf
          If Existe_ID Then
             LineaTexto = LineaTexto & sSQL
             Print #NumFile, LineaTexto
          End If
       End If
       RstSchema.MoveNext
    Loop
    RstSchema.Close
    Close #NumFile
    
    RutaGeneraFile = RutaSysBases & "\TEMP\SCRIPT_CREACION_INDEX.sql"
   'FSeteos.Caption = RutaGeneraFile
    LineaTexto = ""
    NumFile = FreeFile
    Open RutaGeneraFile For Output As #NumFile
    
    Set RstSchema = AdoConexion.OpenSchema(adSchemaTables)
    Do Until RstSchema.EOF
       If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
         'Llenamos la lista de Tablas
          Existe_ID = False
          NombreTabla = RstSchema!TABLE_NAME
          sSQL = "SELECT TOP 1 * " _
               & "FROM " & NombreTabla & " "
          Select_AdoDB AdoTabla, sSQL
          With AdoTabla
               For Idx = 0 To .fields.Count - 1
                   If .fields(Idx).Name = "ID" Then Existe_ID = True
               Next Idx
          End With
          AdoTabla.Close
          sSQL = "ALTER TABLE [T_" & NombreTabla & "] " & vbCrLf _
               & "ADD INDEX ID_" & NombreTabla & " (ModifiedDate);" & vbCrLf _
               & "GO " & vbCrLf
          If Existe_ID Then
             LineaTexto = LineaTexto & sSQL
             Print #NumFile, LineaTexto
          End If
       End If
       RstSchema.MoveNext
    Loop
    RstSchema.Close
    Close #NumFile
    
    AdoConexion.Close
    Progreso_Barra.Mensaje_Box = ""
    Progreso_Final
    MsgBox "Fin del Proceso"
End Sub

Public Sub Eliminar_FN_SP_SQL()
Dim AdoFnSP As ADODB.Recordset
    sSQL = "SELECT ROUTINE_TYPE, ROUTINE_NAME " _
         & "FROM INFORMATION_SCHEMA.ROUTINES " _
         & "WHERE ROUTINE_TYPE IN ('FUNCTION','PROCEDURE') " _
         & "ORDER BY ROUTINE_TYPE, ROUTINE_NAME "
    Select_AdoDB AdoFnSP, sSQL
    With AdoFnSP
     If .RecordCount > 0 Then
         Do While Not .EOF
            sSQL = "IF EXISTS (SELECT name FROM sysobjects WHERE name = '" & .fields("ROUTINE_NAME") & "') "
            If .fields("ROUTINE_TYPE") = "FUNCTION" Then
                sSQL = sSQL & "DROP FUNCTION " & .fields("ROUTINE_NAME") & ";"
            Else
                sSQL = sSQL & "DROP PROCEDURE " & .fields("ROUTINE_NAME") & ";"
            End If
            Ejecutar_SQL_AdoDB sSQL
           .MoveNext
         Loop
     End If
    End With
End Sub

Public Sub Crear_Script_SQL(ProgressBarEstado As ProgressBar, LstStatud As ListBox)
Dim ContFile As Long
Dim NumFile As Long
Dim ArchivoActual As String
Dim ArchivoSQL As String
Dim ListaArchivo() As String
Dim ListaArchivoSPFN() As String
Dim InStrgIni As Integer
Dim InStrgFin As Integer

   Progreso_Barra.Mensaje_Box = "Eliminacion de Funciones y Procedimientos Almacenados"
   ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
   Eliminar_FN_SP_SQL
   
   Progreso_Barra.Mensaje_Box = "Creación de Funciones y Procedimientos Almacenados"
   'Progreso_Iniciar
   ArchivoActual = Dir(RutaSistema & "\BASES\UPDATE_DB\*.sql", vbNormal) 'Recupera la primera entrada.
   NumFile = 1
   Do While ArchivoActual <> ""
      If ArchivoActual <> "." And ArchivoActual <> ".." Then
         ArchivoSQL = ""
         InStrgIni = InStr(ArchivoActual, "UserDefinedFunction")
         If InStrgIni > 1 Then ArchivoSQL = MidStrg(ArchivoActual, 5, InStrgIni - 6)
         InStrgIni = InStr(ArchivoActual, "StoredProcedure")
         If InStrgIni > 1 Then ArchivoSQL = MidStrg(ArchivoActual, 5, InStrgIni - 6)
        'Insertamos el nombre del SP o FN
         If Len(ArchivoSQL) > 1 Then
            ReDim Preserve ListaArchivo(NumFile) As String
            ReDim Preserve ListaArchivoSPFN(NumFile) As String
            
            ListaArchivo(NumFile - 1) = ArchivoSQL
            ListaArchivoSPFN(NumFile - 1) = ArchivoActual
            NumFile = NumFile + 1
         End If
      End If
      ArchivoActual = Dir
   Loop
   
'''   For ContFile = 0 To UBound(ListaArchivo) - 1
'''       sSQL = ""
'''       ArchivoSQL = ListaArchivo(ContFile)
'''       Progreso_Barra.Mensaje_Box = "Eliminando: " & ArchivoSQL
'''       ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''       Select Case MidStrg(ArchivoSQL, 1, 3)
'''         Case "fn_": sSQL = "IF EXISTS (SELECT name FROM sysobjects WHERE name = '" & ArchivoSQL & "') DROP FUNCTION " & ArchivoSQL & ";"
'''         Case "sp_": sSQL = "IF EXISTS (SELECT name FROM sysobjects WHERE name = '" & ArchivoSQL & "') DROP PROCEDURE " & ArchivoSQL & ";"
'''       End Select
'''       If Len(sSQL) > 1 Then Ejecutar_SQL_AdoDB sSQL, True
'''   Next ContFile
   
   RatonReloj
   For ContFile = 0 To UBound(ListaArchivo) - 1
       Progreso_Barra.Mensaje_Box = "Creando " & ListaArchivo(ContFile)
       ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
       Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\" & ListaArchivoSPFN(ContFile)), True
       'MsgBox Progreso_Barra.Mensaje_Box
   Next ContFile
   RatonNormal
   Progreso_Barra.Mensaje_Box = ""
   'Progreso_Final
End Sub

'devuelve un objeto Recordset con los datos de la hoja
Public Function Importar_Excel_AdoDB(ftpLinode As cFTP, LstStatud As ListBox, LstVwFTP As ListView, ByVal PathXls As String) As ADODB.Recordset
Dim obj_Excel     As Object
Dim obj_Workbook  As Object
Dim obj_Worksheet As Object
 
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
   
Dim hoja As String
Dim Carpeta As String
Dim TipoFile As String
Dim cs As String
Dim SQL As String
Dim sSheetName As String
 
   RatonReloj
   EsFileCSV = False
   
   FEsperar.Show
   Imagen_Esperar "Importando el Archivo: " & vbCrLf & PathXls

   TipoFile = Right$(PathXls, Len(PathXls) - InStrRev(PathXls, "."))
   Select Case UCaseStrg(TipoFile)
     Case "TXT"
        Set rs = New ADODB.Recordset
        Set cn = New ADODB.Connection
        
        hoja = Right$(PathXls, Len(PathXls) - InStrRev(PathXls, "\"))
        Carpeta = MidStrg(PathXls, 1, Len(PathXls) - Len(hoja))
       'Cadena de conexión

       'Extensions=asc,csv,tab,txt;HDR=NO;Persist Security Info=False", "", ""
       'cn.open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & Carpeta & ";", "", """"

        cs = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & Carpeta & ";""Extended Properties=""Text;HDR=No;FMT=Delimited(;)"";"
       'Conectamos el archivo de texto o csv
        cn.open cs
       
       'Ejecutamos el recordset pasandole el archivo de texto en la cláusula From
        rs.open "select * from [" & hoja & "]", cn, adOpenStatic, _
        adLockReadOnly, adCmdText
        
        'eliminamos las variables
        Set Importar_Excel_AdoDB = rs
        Set rs = Nothing
        Set cn = Nothing
     Case "CSV"
        Set rs = New ADODB.Recordset
        Set cn = New ADODB.Connection
        EsFileCSV = True
        
        Subir_Archivo_FTP_Linode ftpLinode, LstStatud, LstVwFTP, PathXls
        Subir_Archivo_CSV_SP PathXls
        Eliminar_Archivo_FTP_Linode ftpLinode, LstStatud, LstVwFTP, PathXls
        
        Select_AdoDB rs, "SELECT * FROM Asiento_CSV_" & CodigoUsuario
        Set Importar_Excel_AdoDB = rs
        Set rs = Nothing
     Case "XLS", "XLSX"
          Set rs = New ADODB.Recordset
          rs.CursorLocation = adUseClient
          rs.CursorType = adOpenKeyset
          rs.LockType = adLockBatchOptimistic
         'crea rnueva instancia de Excel
          Set obj_Excel = CreateObject("Excel.Application")
         'Abrir el libro
          Set obj_Workbook = obj_Excel.Workbooks.open(PathXls)
         'referencia la Hoja, por defecto la hoja activa
          If sSheetName = vbNullString Then
             Set obj_Worksheet = obj_Workbook.ActiveSheet
             hoja = obj_Workbook.ActiveSheet.Name
          Else
            'Set obj_Worksheet = obj_Workbook.Sheets(sSheetName)
             hoja = obj_Workbook.Sheets(sSheetName)
          End If
          cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
          hoja = "[" & hoja & "$" & "]"
          rs.open "SELECT * FROM " & hoja, cs
          Set Importar_Excel_AdoDB = rs
          Set rs = Nothing
         'Cerrar libro
          obj_Workbook.Close
         'Cerrar Excel
          obj_Excel.Quit
          Set obj_Workbook = Nothing
          Set obj_Excel = Nothing
          Set obj_Worksheet = Nothing
     Case Else: MsgBox "Formato no permitido"
   End Select
   RatonNormal
   Unload FEsperar
End Function

Public Sub Exportar_AdoDB_Excel(consulta As ADODB.Recordset, Optional nombreHoja As String)
Dim APIExcel As Object
Dim AddLibro As Object
Dim AddHoja  As Object
Dim I As Long
Dim filas As Long
Dim columnas As Long

   If DescripcionEstado = "" Then DescripcionEstado = "NO ESTA PERMITIDO"
   Select Case EstadoEmpresa
     Case "BLOQ", "ONLY", "VEN90", "VEN180", "VEN360", "MAS360", "PRUEBA"
          MsgBox DescripcionEstado, vbOKOnly, "ESTADO DE LA EMPRESA"
     Case Else
          'Mensajes = "Desea Transportar la Consulta a Microsoft Excel "
          'Titulo = "GENERACION DE ARCHIVOS A EXCEL"
          'If BoxMensaje = vbYes Then
             Progreso_Barra.Mensaje_Box = "Exportando a Excel"
             Progreso_Iniciar
        
             RatonReloj
             Progreso_Esperar True
           
             'Creamos objeto excel y nuevo libro y no mostramos el archivo
              Set APIExcel = CreateObject("Excel.Application")
              Set AddLibro = APIExcel.Workbooks.Add
              APIExcel.Visible = False
              
             'Añadimos hoja al libro nuevo y nombramos pestaña
              Set AddHoja = AddLibro.Worksheets(1)
           
             'Damos nombre a la hoja con la que vamos a exportar los datos
             
              'If Len(nombreHoja) > 0 Then AddHoja.Name = MidStrg(nombreHoja, 1, 31) Else
              
              AddHoja.Name = "DiskCover System"
              Progreso_Barra.Mensaje_Box = "DiskCover System"
              Progreso_Esperar True
              
             'Traemos los datos de cabecera de la tabla Access y los pegamos en la hoja excel
              columnas = consulta.fields.Count
              filas = consulta.RecordCount
           
             'Generamos encabezado con colores
              With APIExcel.Range(APIExcel.cells(1, 1), APIExcel.cells(1, columnas))
                  .Font.bold = True
                  .Interior.color = RGB(128, 128, 128)
                  .HorizontalAlignment = 3
                 '.VerticalAlignment = 2
                  .EntireRow.RowHeight = 20
              End With
           
              With APIExcel.Range(APIExcel.cells(2, 1), APIExcel.cells(2, columnas))
                  .Font.bold = True
                  .Interior.color = RGB(168, 168, 0)
                  .HorizontalAlignment = 3
                  .VerticalAlignment = 2
                  .EntireRow.RowHeight = 20
              End With
           
              With APIExcel.Range(APIExcel.cells(2, 1), APIExcel.cells(filas + 2, columnas)).Borders
                  .LineStyle = 1
                  .Weight = 1
                  .ColorIndex = 5
              End With
            
             'Escribimos el encabezado de la consulta
              APIExcel.cells(1, 1) = Empresa & " [" & nombreHoja & "]"
              For I = 0 To columnas - 1
                  APIExcel.cells(2, I + 1) = consulta.fields(I).Name
              Next I
           
             'Pegamos los datos de la tabla en la nueva hoja
              consulta.MoveFirst
              AddHoja.Range("A3").CopyFromRecordset consulta
              
             'Damos formato a las columnas, ajustando contenidos
              With APIExcel.ActiveSheet.cells
                  .Select
                  .EntireColumn.AutoFit
                  .Range("A1").Select
              End With
              
             'Mostramos la hoja
              APIExcel.Visible = True
              RatonNormal
              Progreso_Barra.Mensaje_Box = ""
              Progreso_Final
           'End If
   End Select
End Sub

Public Sub Subir_Archivo_FTP_Linode(ftpLinode As cFTP, LstStatud As ListBox, LstVwFTP As ListView, sOrigen As String)
Dim sDestino As String
Dim FileCSV As String

On Error GoTo error_Handler

   RatonReloj
   If strIPServidor = "mysql.diskcoversystem.com" Then strIPServidor = "db.diskcoversystem.com"
   If strIPServidor = "db.diskcoversystem.com" Then
      With ftpLinode
           Progreso_Barra.Mensaje_Box = "Conectando al servidor"
          .Inicializar MDIFormulario
          'Le establecemos el nombre y Clave del usuario de la cuenta
          .Usuario = ftpUseLinode
          .Password = ftpPwrLinode
          .Puerto = 21
          'Establecesmo el nombre del Servidor FTP
          .servidor = ftpSvrLinode
          'conectamos al servidor FTP. EL label es el control donde mostrar los errores y el estado de la conexión
           If .ConectarFtp(LstStatud) = False Then
               RatonNormal
               MsgBox "Error (" & Err.Number & ") " & Err.Description & vbCrLf & "No se pudo conectar"
               Exit Sub
           End If
          'Mostramos en el label el path del directorio actual donde estamos ubicados en el servidor
           Progreso_Barra.Mensaje_Box = .GetDirectorioActual
          'Le indicamos el ListView donde se listarán los archivos
           Set .ListView = LstVwFTP
           Progreso_Barra.Mensaje_Box = "Buscando directorio en el servidor"
          '----------------------------------------------------------------------------------
          'Realizamos la subida del archivo
          '==================================================================================
           FileCSV = Right$(sOrigen, Len(sOrigen) - InStrRev(sOrigen, "\"))
           sDestino = "/files/" & FileCSV
           Progreso_Barra.Mensaje_Box = "Subiendo: " & FileCSV
          .SubirArchivo sOrigen, sDestino, True
          .Desconectar
      End With
   End If
   Progreso_Barra.Mensaje_Box = "OK"
   Progreso_Esperar True
   RatonNormal
   Exit Sub
error_Handler:
     RatonNormal
     MsgBox Err.Description, vbCritical
End Sub

Public Sub Eliminar_Archivo_FTP_Linode(ftpLinode As cFTP, LstStatud As ListBox, LstVwFTP As ListView, sOrigen As String)
Dim sDestino As String
Dim FileCSV As String

On Error GoTo error_Handler

   RatonReloj
   'MsgBox strIPServidor
   If strIPServidor = "db.diskcoversystem.com" Then
      With ftpLinode
           Progreso_Barra.Mensaje_Box = "Conectando al servidor"
          .Inicializar MDIFormulario
          'Le establecemos el nombre y Clave del usuario de la cuenta
          .Usuario = ftpUseLinode
          .Password = ftpPwrLinode
          .Puerto = 21
          'Establecesmo el nombre del Servidor FTP
          .servidor = ftpSvrLinode
          'conectamos al servidor FTP. EL label es el control donde mostrar los errores y el estado de la conexión
           If .ConectarFtp(LstStatud) = False Then
               RatonNormal
               MsgBox "Error (" & Err.Number & ") " & Err.Description & vbCrLf & "No se pudo conectar"
               Exit Sub
           End If
          'Mostramos en el label el path del directorio actual donde estamos ubicados en el servidor
           Progreso_Barra.Mensaje_Box = .GetDirectorioActual
          'Le indicamos el ListView donde se listarán los archivos
           Set .ListView = LstVwFTP
           Progreso_Barra.Mensaje_Box = "Buscando directorio en el servidor"
          '----------------------------------------------------------------------------------
          'Realizamos la subida del archivo
          '==================================================================================
           FileCSV = Right$(sOrigen, Len(sOrigen) - InStrRev(sOrigen, "\"))
           sDestino = "/files/" & FileCSV
           Progreso_Barra.Mensaje_Box = "Eliminando: " & sOrigen
          .EliminarArchivo sDestino
          .Desconectar
      End With
   End If
   Progreso_Barra.Mensaje_Box = "OK"
   Progreso_Esperar True
   RatonNormal
   Exit Sub
error_Handler:
     RatonNormal
     MsgBox Err.Description, vbCritical
End Sub

