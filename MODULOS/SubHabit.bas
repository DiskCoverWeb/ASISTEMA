Attribute VB_Name = "SubHabit"
Option Explicit

Public Sub Imprimir_Recibo_Hab(DataComp As Adodc, _
                               Optional EsCampoCorto As Boolean)
Dim FormaImp As Byte
Dim SizeLetra As Single
Dim NUsuario As String
On Error GoTo Errorhandler
FormaImp = 1
SizeLetra = 8
RatonReloj
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
'Establecemos Espacios y seteos de impresion
RatonReloj
LetraAnterior = Printer.FontName
CIConLineas = ProcesarSeteos("RE")
EscalaCentimetro 1, TipoTimes, 10
'Iniciamos la impresion
With DataComp.Recordset
 If .RecordCount > 0 Then
    'Detalle de los Abonos
     SumaDebe = 0: SumaHaber = 0: Efectivo = 0
     UltimaLinea = PosLinea
     Printer.FontSize = SetD(9).Tamaño
     PosLinea = SetD(9).PosY
     NUsuario = UCase(.Fields("Nombre_Completo"))
     ConceptoComp = "DE: " & UCase(.Fields("Detalle")) & ", "
     Producto = .Fields("Detalle")
     Do While Not .EOF
        If Producto <> .Fields("Detalle") Then
           ConceptoComp = ConceptoComp & UCase(.Fields("Detalle")) & ", "
           Producto = .Fields("Detalle")
        End If
            If .Fields("Cheq_Dep") <> Ninguno Then
                PrinterFields SetD(10).PosX, PosLinea, .Fields("Banco")
                PrinterFields SetD(12).PosX, PosLinea, .Fields("Cheq_Dep")
                PrinterFields SetD(9).PosX, PosLinea, .Fields("Abono")
                PosLinea = PosLinea + 0.35
            Else
                Efectivo = Efectivo + .Fields("Abono")
            End If
            SumaDebe = SumaDebe + .Fields("Abono")
            SumaHaber = SumaDebe
       .MoveNext
     Loop
     PrinterVariables SetD(8).PosX, SetD(8).PosY, Efectivo
    .MoveFirst
     PosLinea = SetD(13).PosY
     Printer.FontSize = SetD(13).Tamaño
     Do While Not .EOF
        If .Fields("Abono") <> 0 Then
            Producto = .Fields("Producto")
            If .Fields("Mes") <> Ninguno Then Producto = Producto & ": Mes de " & .Fields("Mes")
            PrinterTexto SetD(13).PosX, PosLinea, .Fields("Cta")
            PrinterTexto SetD(14).PosX, PosLinea, Producto
            PrinterFields SetD(16).PosX, PosLinea, .Fields("Abono")
            PosLinea = PosLinea + 0.4
        End If
       .MoveNext
     Loop
     PrinterTexto SetD(14).PosX, PosLinea, "ANTICIPO CLIENTES"
     PrinterVariables SetD(17).PosX, PosLinea, SumaHaber
    .MoveFirst
     ConceptoComp = ConceptoComp & " DEL LOTE No. " & Format(.Fields("Contrato_No"), "#,##0.00")
     If .Fields("T") = "A" Then
         Dibujo = RutaSistema & "\FORMATOS\ANULADO.GIF"
         PrinterPaint Dibujo, 2, UltimaLinea + 1.7, 6, 1.5
     End If
     Printer.FontSize = SetD(18).Tamaño
     PrinterVariables SetD(16).PosX, SetD(18).PosY, SumaDebe
     PrinterVariables SetD(17).PosX, SetD(18).PosY, SumaHaber
     Printer.FontSize = SetD(19).Tamaño
     PrinterVariables SetD(19).PosX, SetD(19).PosY, NUsuario
     PosLinea = PosLinea + 0.4
     MiFecha = .Fields("Fecha")
     Printer.FontBold = False
     Printer.FontSize = SetD(2).Tamaño
     FechaTexto = NombreCiudad & ", " & FechaStrgDias(.Fields("Fecha"))
     PrinterTexto SetD(2).PosX, SetD(2).PosY, FechaTexto
     Printer.FontSize = SetD(3).Tamaño
     PrinterFields SetD(3).PosX, SetD(3).PosY, .Fields("Cliente")
     Printer.FontSize = SetD(4).Tamaño
     PrinterFields SetD(4).PosX, SetD(4).PosY, .Fields("CI_RUC")
     Printer.FontSize = SetD(5).Tamaño
     PrinterLineas SetD(5).PosX, SetD(5).PosY, ConceptoComp, 15
     Printer.FontSize = SetD(6).Tamaño
     PrinterTexto SetD(6).PosX, SetD(6).PosY, Format(SumaHaber, "#,##0.00")
     Printer.FontSize = SetD(7).Tamaño
     PrinterNum SetD(7).PosX, SetD(7).PosY, SumaHaber
     Printer.FontSize = SetD(11).Tamaño
     PrinterTexto SetD(11).PosX, SetD(11).PosY, SetD(11).Encabezado
 End If
End With
MensajeEncabData = ""
Printer.FontName = LetraAnterior
RatonNormal
Printer.EndDoc
RatonNormal
Else
    RatonNormal
End If
Exit Sub
Errorhandler:
    RatonNormal
    ErrorDeImpresion
    Exit Sub
End Sub
