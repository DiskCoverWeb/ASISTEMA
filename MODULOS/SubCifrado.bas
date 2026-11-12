Attribute VB_Name = "RC4Archivos"
Option Explicit

Private Const FIRMA_VALIDACION As String = "[ENC_OK]"

Public Function ExtraerEntre(ByVal Texto As String, _
                             ByVal DelimInicio As String, _
                             ByVal DelimFin As String, _
                             Optional ByVal Comparacion As VbCompareMethod = vbTextCompare) As String
    Dim posInicio As Long
    Dim posFin As Long
    Dim inicioContenido As Long

    ' 1. Localizar el primer delimitador
    posInicio = InStr(1, Texto, DelimInicio, Comparacion)
    If posInicio = 0 Then Exit Function

    ' Calcular dónde comienza el texto real (después del primer delimitador)
    inicioContenido = posInicio + Len(DelimInicio)

    ' 2. Localizar el delimitador final a partir de inicioContenido
    posFin = InStr(inicioContenido, Texto, DelimFin, Comparacion)
    If posFin = 0 Then Exit Function

    ' 3. Extraer la subcadena
    ExtraerEntre = Trim$(Mid$(Texto, inicioContenido, posFin - inicioContenido))
End Function

' Cifra el texto agregando la firma de verificación al inicio
Public Sub CifrarCadenaYGuardarArchivo(ByVal sTextoOriginal As String, ByVal sRutaDestino As String, ByVal sClave As String)
    Dim sTextoConFirma As String
    Dim sCifradoHex As String
    
    ' Agregamos el identificador de control al inicio del contenido
    sTextoConFirma = FIRMA_VALIDACION & sTextoOriginal
    
    sCifradoHex = RC4_Cifrar(sTextoConFirma, sClave)
    Guardar_Archivo_Texto_Cifrado sRutaDestino, sCifradoHex
End Sub

' Descifra y valida si la clave fue la correcta
Public Function DescifrarArchivoATexto(ByVal sRutaCifrada As String, ByVal sClave As String) As String
    Dim sCifradoHex As String
    Dim sTextoDescifrado As String
    
    If Dir$(sRutaCifrada) = "" Then
        DescifrarArchivoATexto = ""
        Exit Function
    End If
    
    sCifradoHex = Leer_Archivo_Texto_Cifrado(sRutaCifrada)
    sTextoDescifrado = RC4_Descifrar(Trim$(sCifradoHex), sClave)
    
    ' Comprobamos si la firma de control coincide
    If Left$(sTextoDescifrado, Len(FIRMA_VALIDACION)) = FIRMA_VALIDACION Then
        ' Clave correcta: extrae el texto original limpio
        DescifrarArchivoATexto = Mid$(sTextoDescifrado, Len(FIRMA_VALIDACION) + 1)
    Else
        ' Clave incorrecta: devuelve vacío para que el formulario controle el aviso
        DescifrarArchivoATexto = ""
    End If
End Function

' =========================================================================
' MÃ³dulo de Cifrado y Descifrado RC4 con E/S de Archivos en VB6
' =========================================================================

Public Function Leer_Archivo_Texto_Cifrado(ByVal sRuta As String) As String
    Dim numArchivo As Integer
    Dim contenido As String
    
    If Dir$(sRuta) = "" Then
        Err.Raise vbObjectError + 513, "Leer_Archivo_Texto_Cifrado", "El archivo no existe: " & sRuta
    End If
    
    numArchivo = FreeFile
    Open sRuta For Binary Access Read As #numArchivo
        If LOF(numArchivo) > 0 Then
            contenido = Space$(LOF(numArchivo))
            Get #numArchivo, , contenido
        Else
            contenido = ""
        End If
    Close #numArchivo
    
    Leer_Archivo_Texto_Cifrado = contenido
End Function

Public Sub Guardar_Archivo_Texto_Cifrado(ByVal sRuta As String, ByVal sContenido As String)
    Dim numArchivo As Integer
    
    If Dir$(sRuta) <> "" Then Kill sRuta
    
    numArchivo = FreeFile
    Open sRuta For Output As #numArchivo
        Print #numArchivo, sContenido;
    Close #numArchivo
End Sub

Public Function RC4_Cifrar(ByVal Texto As String, ByVal Clave As String) As String
    Dim sBox(0 To 255) As Integer
    Dim keyBox(0 To 255) As Integer
    Dim temp As Integer
    Dim i As Long, j As Long, k As Long
    Dim lenKey As Long, lenTexto As Long
    Dim c As Long, cipherBy As Integer
    Dim res() As Byte
    
    lenKey = Len(Clave)
    lenTexto = Len(Texto)
    
    If lenKey = 0 Or lenTexto = 0 Then
        RC4_Cifrar = ""
        Exit Function
    End If
    
    For i = 0 To 255
        sBox(i) = i
        keyBox(i) = Asc(Mid$(Clave, (i Mod lenKey) + 1, 1))
    Next i
    
    j = 0
    For i = 0 To 255
        j = (j + sBox(i) + keyBox(i)) Mod 256
        temp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = temp
    Next i
    
    i = 0: j = 0
    ReDim res(1 To lenTexto)
    
    For k = 1 To lenTexto
        i = (i + 1) Mod 256
        j = (j + sBox(i)) Mod 256
        temp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = temp
        c = (sBox(i) + sBox(j)) Mod 256
        cipherBy = sBox(c)
        res(k) = Asc(Mid$(Texto, k, 1)) Xor cipherBy
    Next k
    
    RC4_Cifrar = BytesToHex(res)
End Function

Public Function RC4_Descifrar(ByVal TextoHex As String, ByVal Clave As String) As String
    Dim sBox(0 To 255) As Integer
    Dim keyBox(0 To 255) As Integer
    Dim temp As Integer
    Dim i As Long, j As Long, k As Long
    Dim lenKey As Long, lenTexto As Long
    Dim c As Long, cipherBy As Integer
    Dim res As String
    Dim textoBytes As String
    
    TextoHex = Trim$(TextoHex)
    If Len(TextoHex) = 0 Or Len(Clave) = 0 Then
        RC4_Descifrar = ""
        Exit Function
    End If
    
    textoBytes = HexToBytes(TextoHex)
    lenKey = Len(Clave)
    lenTexto = Len(textoBytes)
    
    For i = 0 To 255
        sBox(i) = i
        keyBox(i) = Asc(Mid$(Clave, (i Mod lenKey) + 1, 1))
    Next i
    
    j = 0
    For i = 0 To 255
        j = (j + sBox(i) + keyBox(i)) Mod 256
        temp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = temp
    Next i
    
    i = 0: j = 0
    res = Space$(lenTexto)
    
    For k = 1 To lenTexto
        i = (i + 1) Mod 256
        j = (j + sBox(i)) Mod 256
        temp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = temp
        c = (sBox(i) + sBox(j)) Mod 256
        cipherBy = sBox(c)
        Mid$(res, k, 1) = Chr$(Asc(Mid$(textoBytes, k, 1)) Xor cipherBy)
    Next k
    
    RC4_Descifrar = res
End Function

Private Function BytesToHex(ByRef arrBytes() As Byte) As String
    Dim i As Long, sHex As String
    For i = LBound(arrBytes) To UBound(arrBytes)
        sHex = Hex$(arrBytes(i))
        If Len(sHex) = 1 Then sHex = "0" & sHex
        BytesToHex = BytesToHex & sHex
    Next i
End Function

Private Function HexToBytes(ByVal sHex As String) As String
    Dim i As Long, strOut As String
    On Error Resume Next
    For i = 1 To Len(sHex) Step 2
        strOut = strOut & Chr$(Val("&H" & Mid$(sHex, i, 2)))
    Next i
    HexToBytes = strOut
End Function
