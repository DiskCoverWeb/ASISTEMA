Attribute VB_Name = "SubPrintersPDF"
Option Explicit

Public Function Encabezado_PDF(Nombre_Archivo As String, _
                               Titulo_Archivo As String, _
                               NombreTipoDeLetra As String, _
                               VerDocumento As Boolean, _
                               Optional AImpresora As Boolean) As Single
Dim PosLineaT As Single
Dim PosLinIni As Single
Dim TPosLinea As Single
   'Datos Iniciales
    HoraSistema = Format$(Time, FormatoTimes)
    MiHora = FechaSistema & " - " & HoraSistema
''    Progreso_Barra.Mensaje_Box = "Documento RIDE Electronico"
''    Progreso_Esperar
   'Generamos el documento
    Nombre_Archivo = Nombre_Archivo & "_" & Replace(FechaSistema, "/", "-") & "_" & CodigoUsuario
    If AImpresora Then tPrint.TipoImpresion = Es_Printer Else tPrint.TipoImpresion = Es_PDF
    tPrint.NombreArchivo = Nombre_Archivo
    tPrint.TituloArchivo = Titulo_Archivo
    tPrint.TipoLetra = NombreTipoDeLetra
    tPrint.OrientacionPagina = Orientacion_Pagina
    tPrint.PaginaA4 = True
    tPrint.EsCampoCorto = False
    tPrint.VerDocumento = VerDocumento
    
    Set cPrint = New cImpresion
    cPrint.iniciaImpresion
    
    cPrint.vPaginaNo = 1
   'NombreTipoDeLetra
    cPrint.printEncabezado 1.1, 1, NombreTipoDeLetra
   'Fin de Impresion del Encabezado del Documento PDF
    Encabezado_PDF = PosLinea
End Function


Public Sub Encabezado_Rol()
Dim Ancho_Maximo As Single
 PosLinea = 1
 Ancho_Maximo = cPrint.dAnchoPapel - 0.5
 cPrint.printImagen LogoTipo, 1, PosLinea, 4.5, 2
 RutaDestino = RutaSistema & "\LOGOS\DiskCover.gif"
 cPrint.printImagen RutaDestino, Ancho_Maximo - 1.8, PosLinea, 1.8, 0.6
 cPrint.letraTipo TipoHelvetica, 6
 cPrint.tipoNegrilla = True
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea, "Hora:"
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea + 0.3, "Pagina No."
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea + 0.6, "Fecha:"
 cPrint.printTexto Ancho_Maximo - 4.5, PosLinea + 0.9, "Usuario:"
 cPrint.tipoNegrilla = False
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea, Format(Time, "hh:mm:ss")
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea + 0.3, Format(Pagina, "0000")
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea + 0.6, FechaStrgDias(date)
 cPrint.printTexto Ancho_Maximo - 3.2, PosLinea + 0.9, ULCase(NombreUsuario)
 cPrint.letraTipo TipoTimes
 cPrint.tipoNegrilla = True
 cPrint.PorteDeLetra = 14
 If UCase$(RazonSocial) = UCase$(NombreComercial) Then
    cPrint.printTexto 1, PosLinea, UCase$(RazonSocial), 14, "C", Ancho_Maximo
 Else
    cPrint.printTexto 1, PosLinea, UCase$(RazonSocial), 14, "C", Ancho_Maximo
    cPrint.printTexto 1, PosLinea + 0.5, UCase$(NombreComercial), 14, "C", Ancho_Maximo
 End If
 PosLinea = PosLinea + 0.8
 cPrint.PorteDeLetra = 9
 cPrint.tipoNegrilla = False
 cPrint.printTexto 1, PosLinea, ULCase(Direccion) & ". Teléfono: " & Telefono1, 9, "C", Ancho_Maximo
 PosLinea = PosLinea + 0.45
 cPrint.PorteDeLetra = 12
 cPrint.tipoNegrilla = True
 cPrint.printTexto 1, PosLinea, MensajeEncabData, 12, "C", Ancho_Maximo
 cPrint.tipoNegrilla = False
 cPrint.PorteDeLetra = 8
 Pagina = Pagina + 1
 cPrint.letraTipo TipoHelvetica
 PosLinea = PosLinea + 0.4
End Sub

Public Sub Tipo_Balance_PDF(sSQL_Ext As String, _
                            NombreTipoDeLetra As String, _
                            FechaDesde As String, _
                            FechaHasta As String, _
                            TipoBalance As String, _
                            TipoCC As String, _
                            TipoPyGCC As String)
Dim AdoDBCat As ADODB.Recordset
Dim tipoDeLetra As String
Dim Cod_Aux As String
Dim Cod_Sup As String
Dim CantEsp As Single
Dim CmCadena As Single
     'TipoArial / TipoVerdana / TipoHelvetica
      Titulo = "CONFIRMACION DE IMPRESION"
      Mensajes = "Quiere Enviar a PDF el Documento"
      If BoxMensaje = vbYes Then
         RatonReloj
         FEsperar.Show
        'sSQL = SQL_Tipo_Balance(TipoBalance)
        'MsgBox sSQL_Ext
         Select_AdoDB AdoDBCat, sSQL_Ext
    
        'Encabezado Balance
         Imagen_Esperar
         tipoDeLetra = TipoCourierNew
         SQLMsg1 = "D E L  " & FechaDesde & "  A L  " & FechaHasta
         If TipoPyGCC <> Ninguno Then SQLMsg2 = "CENTRO DE COSTO POR: " & TipoCC Else SQLMsg2 = ""

         PosLinea = Encabezado_PDF(TipoBalance, MensajeEncabData, NombreTipoDeLetra, True)
         SQLMsg3 = ""
         
         cPrint.PorteDeLetra = 8
         MensajeEncabData = ""
         SQLMsg1 = ""
         With AdoDBCat
          If .RecordCount > 0 Then
              Imagen_Esperar
              Encabezado_Campos_Balances_PDF
'''              If MidStrg(.Fields("Codigo"), 1, 2) = "01" Then
'''                 cPrint.printTexto 2.5, PosLinea, "Activos"
'''                 PosLinea = PosLinea + 0.4
'''              End If
              Do While Not .EOF
                 cPrint.tipoNegrilla = False
                 Imagen_Esperar "Procesando: " & .Fields("Codigo")
                 'If Len(.Fields("Codigo")) = 2 Or MidStrg(.Fields("Codigo"), 4, 7) = "9999999" Then cPrint.tipoNegrilla = True
                 CantEsp = 1.5 + Redondear((Niveles(.Fields("Codigo")) - 1) / 5, 2)
                 If .Fields("DG") = "G" Then
                     cPrint.tipoNegrilla = True
                 Else
                     cPrint.printTexto 2.6, PosLinea, MidStrg(.Fields("Codigo"), 6, 7)
                     CantEsp = 3.9
                 End If
                 'MsgBox CantEsp & " - " & .Fields("Codigo_Ext") & " - " & .Fields("Cuenta")
                 Cuenta = .Fields("Cuenta")
                 CmCadena = cPrint.anchoTexto(Cuenta)
                 Do While CmCadena > 6.2
                    Cuenta = MidStrg(Cuenta, 1, Len(Cuenta) - 1)
                    CmCadena = cPrint.anchoTexto(Cuenta)
                 Loop
                 cPrint.printTexto CantEsp, PosLinea, Cuenta
                 cPrint.printFields 10, PosLinea, .Fields("Saldo_Anterior"), 8, , , 2
                 cPrint.printFields 12, PosLinea, .Fields("Debitos"), 8, , , 2
                 cPrint.printFields 14, PosLinea, .Fields("Creditos"), 8, , , 2
                 cPrint.printFields 16, PosLinea, .Fields("Saldo_Mes"), 8, , , 2
                 cPrint.printFields 18, PosLinea, .Fields("Saldo_Total"), 8, , , 2
                 cPrint.tipoNegrilla = False
                 PosLinea = PosLinea + 0.4
                 'If MidStrg(.Fields("Codigo"), 4, 7) = "9999999" Then PosLinea = PosLinea + 0.2
                 If PosLinea > 27 Then
                    cPrint.paginaNueva
                    Imagen_Esperar
                    cPrint.printEncabezado 1.1, 1, NombreTipoDeLetra
                    cPrint.PorteDeLetra = 8
                    Encabezado_Campos_Balances_PDF
                    Imagen_Esperar
                 End If
                .MoveNext
              Loop
              PosLinea = PosLinea + 0.4
          End If
         End With
         AdoDBCat.Close
         RatonNormal
         cPrint.finalizaImpresion
         Unload FEsperar
      End If
End Sub

Public Sub Encabezado_Campos_Balances_PDF()
    cPrint.colorDeLetra = Negro
    cPrint.tipoNegrilla = True
    cPrint.printTexto 1.5, PosLinea, "Codigo"
    cPrint.printTexto 3.5, PosLinea, "Nombre_Cuenta"
    cPrint.printTexto 10.4, PosLinea, "Saldo_Anterior"
    cPrint.printTexto 13.3, PosLinea, "Debitos"
    cPrint.printTexto 15.2, PosLinea, "Creditos"
    cPrint.printTexto 16.8, PosLinea, "Saldo_Mes"
    cPrint.printTexto 18.7, PosLinea, "Saldo_Total"
    PosLinea = PosLinea + 0.4
    cPrint.printLinea 1, PosLinea, 20, PosLinea
    PosLinea = PosLinea + 0.05
End Sub

Public Sub PDF_Analitico_Mensual(AdoAnaliticoMensual As Adodc, _
                                 Encabezado As String, vFechaFinal As String)
Dim PorteDeLetra As Integer
Dim tipoDeLetra As String

   RatonReloj
   Bandera = False
   If Month(vFechaFinal) >= 6 Then Orientacion_Pagina = 2 Else Orientacion_Pagina = 1
   PorteDeLetra = 6
   PosLinea = 1.1
   tipoDeLetra = TipoArial 'TipoHelvetica
   MensajeEncabData = Encabezado
  'Generamos el nombre del documento
   tPrint.TipoImpresion = Es_PDF
   tPrint.NombreArchivo = "Analitico_Mensual_" & NumEmpresa & "_" & CodigoUsuario
   tPrint.TituloArchivo = "Estado Analitico Mensual"
   tPrint.TipoLetra = tipoDeLetra ' TipoHelvetica 'TipoArial
   tPrint.PorteLetra = PorteDeLetra
   tPrint.OrientacionPagina = Orientacion_Pagina
   tPrint.PaginaA4 = True
   tPrint.EsCampoCorto = False
   tPrint.VerDocumento = True
  'Empezamos a llenar el contenido
   Set cPrint = New cImpresion
   cPrint.iniciaImpresion
  'MsgBox cPrint.dAltoPapel & vbCrLf & cPrint.dAnchoPapel & vbCrLf & LimiteAlto & vbCrLf & LimiteAncho
   Progreso_Barra.Mensaje_Box = "Analitico Mensual " & Codigo
   Progreso_Iniciar
   With AdoAnaliticoMensual.Recordset
    If .RecordCount > 0 Then
        Progreso_Barra.Valor_Maximo = .RecordCount
        cPrint.anchoRegistro 1.2, AdoAnaliticoMensual, True
        cPrint.printEncabezado 1.2, PosLinea, tipoDeLetra
        cPrint.printAllFields 1.2, PosLinea, AdoAnaliticoMensual, False, True, PorteDeLetra, , 2
        PosLinea = PosLinea + 0.4
        Do While Not .EOF
           Progreso_Barra.Mensaje_Box = "Generando PDF de Analitico Mensual"
           Progreso_Esperar
           If .Fields("DG") = "G" Then cPrint.tipoNegrilla = True
           If .Fields("TC") <> "N" Then cPrint.tipoItalica = True
           If .Fields("TC") = "C" Then cPrint.tipoSubrayado = True
           If .Fields("TC") = "P" Then cPrint.tipoSubrayado = True
           If .Fields("TC") = "I" Then cPrint.tipoSubrayado = True
           If .Fields("TC") = "G" Then cPrint.tipoSubrayado = True
           If .Fields("TC") = "CC" Then cPrint.tipoSubrayado = True
           If .Fields("Cta") = "  " Then
               cPrint.tipoNegrilla = False
               cPrint.tipoItalica = False
               cPrint.tipoSubrayado = False
           End If
           If .Fields("Cta") = "(+/-)" Then
               PosLinea = PosLinea + 0.1
               cPrint.printLinea 1, PosLinea, LimiteAncho, PosLinea
               PosLinea = PosLinea + 0.05
           End If
           cPrint.printAllFields 1.2, PosLinea, AdoAnaliticoMensual, False, False, PorteDeLetra, , 2
           cPrint.tipoNegrilla = False
           cPrint.tipoItalica = False
           cPrint.tipoSubrayado = False
           If PosLinea > LimiteAlto Then
              cPrint.paginaNueva
              PosLinea = 1.2
              cPrint.printEncabezado 1.2, PosLinea, tipoDeLetra
              cPrint.printAllFields 1.2, PosLinea, AdoAnaliticoMensual, False, True, PorteDeLetra, , 2
              PosLinea = PosLinea + 0.4
           Else
              PosLinea = PosLinea + 0.35
           End If
           
          'MsgBox Progreso_Barra.Incremento & "... " & .Fields("Cta")
          .MoveNext
        Loop
       .MoveFirst
    End If
   End With
   cPrint.finalizaImpresion
   MensajeEncabData = ""
   RatonNormal
   Progreso_Final
End Sub

