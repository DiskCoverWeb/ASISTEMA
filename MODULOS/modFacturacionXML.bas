Attribute VB_Name = "modFacturacionXML"
Option Explicit

' APIs de Windows para esperar la ejecución del conversor PDF
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const SYNCHRONIZE As Long = &H100000
Private Const INFINITE As Long = &HFFFFFFFF

' =========================================================================
' RUTINA PRINCIPAL: Lee el XML, extrae los datos y crea el PDF
' =========================================================================
Public Function GenerarPDF_Desde_XML(ByVal rutaXML As String, ByVal rutaPDF As String) As Boolean
    On Error GoTo ErrHandler

    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.30")
    
    ' === OBLIGATORIO EN MSXML 3.0 ===
    xmlDoc.async = False
    xmlDoc.setProperty "SelectionLanguage", "XPath"
    ' ================================
    
    If xmlDoc.Load(rutaXML) Then
        ' Todo lo demás (.selectSingleNode, .selectNodes, .Text) funciona exactamente igual
        Dim sNumero As String
        sNumero = xmlDoc.selectSingleNode("//*[local-name()='ID']").Text
        MsgBox "Comprobante: " & sNumero
    Else
        MsgBox "Error al cargar el XML: " & xmlDoc.parseError.reason, vbCritical, "Error XML"
        GenerarPDF_Desde_XML = False
        Exit Function
    End If
        
    ' -------------------------------------------------------------
    ' 1. EXTRAER DATOS DEL COMPROBANTE (Soporta UBL 2.0 / 2.1)
    ' -------------------------------------------------------------
    Dim sTipoDoc As String
    Dim sNumeroDoc As String
    Dim sFechaEmision As String
    Dim sMoneda As String
    
    ' Datos Emisor
    Dim sEmisorRuc As String
    Dim sEmisorRazon As String
    Dim sEmisorDireccion As String
    
    ' Datos Cliente
    Dim sClienteDoc As String
    Dim sClienteRazon As String
    
    ' Totales y Firma
    Dim sTotal As String
    Dim sSubTotal As String
    Dim sIgv As String
    Dim sHashDigest As String
    
    ' Obtener valores usando helper para ignorar prefijos de namespace (cac:, cbc:, etc.)
    sNumeroDoc = ObtenerValorNodo(xmlDoc, "ID")
    sFechaEmision = ObtenerValorNodo(xmlDoc, "IssueDate")
    sMoneda = ObtenerValorNodo(xmlDoc, "DocumentCurrencyCode")
    
    ' Emisor
    sEmisorRuc = ObtenerValorNodo(xmlDoc, "AccountingSupplierParty//CompanyID")
    sEmisorRazon = ObtenerValorNodo(xmlDoc, "AccountingSupplierParty//RegistrationName")
    sEmisorDireccion = ObtenerValorNodo(xmlDoc, "AccountingSupplierParty//Line")
    
    ' Cliente / Receptor
    sClienteDoc = ObtenerValorNodo(xmlDoc, "AccountingCustomerParty//CompanyID")
    If sClienteDoc = "" Then sClienteDoc = ObtenerValorNodo(xmlDoc, "AccountingCustomerParty//CustomerAssignedAccountID")
    sClienteRazon = ObtenerValorNodo(xmlDoc, "AccountingCustomerParty//RegistrationName")
    
    ' Totales
    sTotal = ObtenerValorNodo(xmlDoc, "LegalMonetaryTotal/PayableAmount")
    If sTotal = "" Then sTotal = ObtenerValorNodo(xmlDoc, "LegalMonetaryTotal/TaxInclusiveAmount")
    
    ' DigestValue (Código Hash de la firma digital del XML)
    sHashDigest = ObtenerValorNodo(xmlDoc, "DigestValue")
    
    ' -------------------------------------------------------------
    ' 2. EXTRAER LOS ÍTEMS / LÍNEAS DEL COMPROBANTE
    ' -------------------------------------------------------------
    Dim nodeListItems As Object
    Dim nodoItem As Object
    Dim sFilasHTML As String
    Dim itemIdx As Long
    
    ' Busca las líneas (InvoiceLine, CreditNoteLine, DebitNoteLine)
    Set nodeListItems = xmlDoc.selectNodes("//*[local-name()='InvoiceLine' or local-name()='CreditNoteLine' or local-name()='DebitNoteLine']")
    
    itemIdx = 1
    For Each nodoItem In nodeListItems
        Dim sCant As String, sDesc As String, sPreUni As String, sTotalLin As String
        
        sCant = ObtenerValorSubNodo(nodoItem, "InvoicedQuantity")
        If sCant = "" Then sCant = ObtenerValorSubNodo(nodoItem, "CreditedQuantity")
        If sCant = "" Then sCant = ObtenerValorSubNodo(nodoItem, "DebitedQuantity")
        
        sDesc = ObtenerValorSubNodo(nodoItem, "Item/Description")
        sPreUni = ObtenerValorSubNodo(nodoItem, "Price/PriceAmount")
        sTotalLin = ObtenerValorSubNodo(nodoItem, "LineExtensionAmount")
        
        sFilasHTML = sFilasHTML & "<tr>" & _
                     "<td align='center'>" & itemIdx & "</td>" & _
                     "<td align='center'>" & sCant & "</td>" & _
                     "<td>" & sDesc & "</td>" & _
                     "<td align='right'>" & FormatNumber(Val(sPreUni), 2) & "</td>" & _
                     "<td align='right'>" & FormatNumber(Val(sTotalLin), 2) & "</td>" & _
                     "</tr>"
        itemIdx = itemIdx + 1
    Next nodoItem
    
    ' -------------------------------------------------------------
    ' 3. ARMAR EL DISEÑO HTML
    ' -------------------------------------------------------------
    Dim sHtml As String
    sHtml = "<!DOCTYPE html><html><head><meta charset='utf-8'>" & _
            "<style>" & _
            "  body { font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #333; margin: 15px; }" & _
            "  .box-ruc { border: 2px solid #000; border-radius: 6px; text-align: center; width: 230px; float: right; padding: 10px; font-weight: bold; font-size: 13px; }" & _
            "  .tabla-datos { width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 10px; }" & _
            "  .tabla-datos td { padding: 4px; }" & _
            "  .tabla-items { width: 100%; border-collapse: collapse; margin-top: 10px; }" & _
            "  .tabla-items th { background: #333; color: #fff; padding: 6px; border: 1px solid #333; font-size: 10px; }" & _
            "  .tabla-items td { padding: 5px; border: 1px solid #ccc; font-size: 10px; }" & _
            "  .totales { float: right; width: 220px; margin-top: 10px; }" & _
            "  .totales td { padding: 3px; }" & _
            "  .hash-box { margin-top: 25px; font-family: monospace; font-size: 9px; color: #666; border: 1px dashed #aaa; padding: 5px; width: 60%; float: left; }" & _
            "</style></head><body>"
            
    ' Encabezado
    sHtml = sHtml & "<div style='width: 100%;'>"
    sHtml = sHtml & "  <div class='box-ruc'>R.U.C. " & sEmisorRuc & "<br>COMPROBANTE ELECTRÓNICO<br>" & sNumeroDoc & "</div>"
    sHtml = sHtml & "  <h2 style='margin:0;'>" & sEmisorRazon & "</h2>"
    sHtml = sHtml & "  <p style='margin:3px 0;'>" & sEmisorDireccion & "</p>"
    sHtml = sHtml & "  <div style='clear:both;'></div>"
    sHtml = sHtml & "</div>"
    
    ' Datos del Cliente
    sHtml = sHtml & "<table class='tabla-datos'>"
    sHtml = sHtml & "  <tr><td width='15%'><b>Señor(es):</b></td><td>" & sClienteRazon & "</td><td width='15%'><b>Fecha Emisión:</b></td><td>" & sFechaEmision & "</td></tr>"
    sHtml = sHtml & "  <tr><td><b>RUC/DNI:</b></td><td>" & sClienteDoc & "</td><td><b>Moneda:</b></td><td>" & sMoneda & "</td></tr>"
    sHtml = sHtml & "</table>"
    
    ' Detalle de Ítems
    sHtml = sHtml & "<table class='tabla-items'>"
    sHtml = sHtml & "  <thead><tr><th width='5%'>Item</th><th width='10%'>Cant.</th><th>Descripción</th><th width='15%'>V. Unit</th><th width='15%'>Importe</th></tr></thead>"
    sHtml = sHtml & "  <tbody>" & sFilasHTML & "</tbody>"
    sHtml = sHtml & "</table>"
    
    ' Hash y Totales
    sHtml = sHtml & "<div class='hash-box'><b>Resumen Digital (Hash / DigestValue):</b><br>" & sHashDigest & "</div>"
    sHtml = sHtml & "<div class='totales'>"
    sHtml = sHtml & "  <table width='100%'>"
    sHtml = sHtml & "    <tr><td><b>Importe Total:</b></td><td align='right'><b>" & sMoneda & " " & FormatNumber(Val(sTotal), 2) & "</b></td></tr>"
    sHtml = sHtml & "  </table>"
    sHtml = sHtml & "</div>"
    
    sHtml = sHtml & "</body></html>"
    
    ' -------------------------------------------------------------
    ' 4. GUARDAR HTML TEMPORAL Y GENERAR EL PDF
    ' -------------------------------------------------------------
    Dim sRutaTempHtml As String
    Dim nFile As Integer
    sRutaTempHtml = App.Path & "\temp_" & sNumeroDoc & ".html"
    
    nFile = FreeFile
    Open sRutaTempHtml For Output As #nFile
    Print #nFile, sHtml
    Close #nFile
    
    ' Convertir HTML a PDF (mediante Microsoft Edge Headless nativo en Windows 10/11 o wkhtmltopdf)
    GenerarPDF_Desde_HTML sRutaTempHtml, rutaPDF
    
    ' Limpiar archivo temporal
    On Error Resume Next
    Kill sRutaTempHtml
    
    GenerarPDF_Desde_XML = (Dir(rutaPDF) <> "")
    Exit Function

ErrHandler:
    GenerarPDF_Desde_XML = False
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Function

' =========================================================================
' FUNCIONES AUXILIARES PARA LEER NODOS XML INDEPENDIENTE DEL NAMESPACE
' =========================================================================
Private Function ObtenerValorNodo(ByRef xmlDoc As Object, ByVal sNodeName As String) As String
    Dim nodo As Object
    ' Usa local-name() para leer tags sin importar el prefijo cac:, cbc:, o si no tiene
    If InStr(sNodeName, "//") > 0 Or InStr(sNodeName, "/") > 0 Then
        Set nodo = xmlDoc.selectSingleNode("//*[local-name()='" & sNodeName & "']")
        If nodo Is Nothing Then
            ' Intento de búsqueda directa por XPath simplificado
            Set nodo = xmlDoc.selectSingleNode("//" & sNodeName)
        End If
    Else
        Set nodo = xmlDoc.selectSingleNode("//*[local-name()='" & sNodeName & "']")
    End If
    
    If Not nodo Is Nothing Then
        ObtenerValorNodo = Trim(nodo.Text)
    Else
        ObtenerValorNodo = ""
    End If
End Function

Private Function ObtenerValorSubNodo(ByRef parentNode As Object, ByVal sPath As String) As String
    Dim nodo As Object
    Set nodo = parentNode.selectSingleNode(".//*[local-name()='" & sPath & "']")
    If nodo Is Nothing Then
        Set nodo = parentNode.selectSingleNode(".//" & sPath)
    End If
    
    If Not nodo Is Nothing Then
        ObtenerValorSubNodo = Trim(nodo.Text)
    Else
        ObtenerValorSubNodo = ""
    End If
End Function

' =========================================================================
' MOTOR DE IMPRESIÓN A PDF (Silencioso y sin interfaz)
' =========================================================================
Private Sub GenerarPDF_Desde_HTML(ByVal sHtmlPath As String, ByVal sPdfPath As String)
    Dim cmd As String
    Dim procId As Long, hProcess As Long
    
    ' Si tienes wkhtmltopdf.exe en la carpeta de la app lo usa; si no, usa Microsoft Edge
    If Dir(App.Path & "\wkhtmltopdf.exe") <> "" Then
        cmd = """" & App.Path & "\wkhtmltopdf.exe"" -q --page-size A4 """ & sHtmlPath & """ """ & sPdfPath & """"
    Else
        ' Microsoft Edge viene instalado por defecto en todos los Windows modernos
        cmd = "msedge.exe --headless --disable-gpu --run-all-compositor-stages-before-draw --print-to-pdf=""" & sPdfPath & """ """ & sHtmlPath & """"
    End If
    
    procId = Shell(cmd, vbHide)
    If procId <> 0 Then
        hProcess = OpenProcess(SYNCHRONIZE, False, procId)
        If hProcess <> 0 Then
            WaitForSingleObject hProcess, INFINITE
            CloseHandle hProcess
        End If
    End If
End Sub

