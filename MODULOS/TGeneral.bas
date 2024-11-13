Attribute VB_Name = "TTipos"
Option Explicit

'Type para el Rango
Type T_Rango
     NumFila1 As Long
     NumFila2 As Long
     NumCol1 As Long
     NumCol2 As Long
End Type
'-------------------------------------
' TIPOS PERSONALIZADOS
'-------------------------------------
Public Enum TipoBusqueda
       X_CI_RUC = 1
       X_Beneficiario = 2
       X_Grupo = 3
End Enum
'-------------------------------------
' TIPOS PRINCIPALES
'-------------------------------------
Type POINTS
     X As Long
     y As Long
End Type
'-------------------------------------
Type Nodo_Arbol
     Item_Nodo  As String
     Codigo_Aux As String
     Valor      As Currency
     Eliminar   As Boolean
End Type
'-------------------------------------
 Type NombreTablas
      Nombre   As String
      Cantidad As Integer
 End Type
'-------------------------------------
 Type EncabezadoReporte
      MsgTitulo     As String
      MsgObjetivo   As String
      MsgConcepto   As String
      TextoObjetivo As String
      TextoConcepto As String
 End Type
'-------------------------------------
 Type Campos_Tabla
      Campo      As String
      Ancho      As Long
      Tipo       As Long
      Si_ID      As Boolean
      Si_Periodo As Boolean
      Valor      As Variant
 End Type
'-------------------------------------
 Type Campos_Decimal
      Campo      As String
      CantDec    As Byte
      AnchoCampo As Single
      SumaTotal  As Variant
 End Type
'-------------------------------------
 Type Crear_Tablas
      Campo      As String
      TipoAdodc  As String
      TipoSQL    As String
      TipoAccess As String
      ErrorCampo As Boolean
      LargoCampo As Long
 End Type
'-------------------------------------
 Type Campos_Rol
      DC As String
      Lst As Boolean
      Campo As String
      Detalle As String
      Codigo As String
 End Type
'-------------------------------------
 Type Grafico
      Titulo  As String
      TituloX As String
      TituloY As String
 End Type
'-------------------------------------
 Type Comprobantes
      T                  As String
      TP                 As String
      Fecha              As String
      CodigoB            As String
      CodigoDr           As String
      Beneficiario       As String
      RUC_CI             As String
      TD                 As String
      Telefono           As String
      Direccion          As String
      Email              As String
      AgenteRetencion    As String
      MicroEmpresa       As String
      Estado             As String
      Concepto           As String
      Usuario            As String
      Autorizado         As String
      Item               As String
      Ctas_Modificar     As String
      CodigoInvModificar As String
      Grupo              As String
      TipoContribuyente  As String
      Cheque             As String
      Cta_Banco          As String
      Serie_R            As String
      Autorizacion_R     As String
      Serie_LC           As String
      Autorizacion_LC    As String
      Cotizacion         As Single
      Efectivo           As Currency
      Total_Banco        As Currency
      Monto_Total        As Currency
      Numero             As Long
      Retencion          As Long
      T_No               As Byte
      RetNueva           As Boolean
      RetSecuencial      As Boolean
 End Type
'-------------------------------------
 Type Bancos
      Cta_Banco As String
      Fecha     As String
      TP        As String
      Numero    As Long
      Cheq_Dep  As String
      Valor     As Currency
      M_E       As Boolean
 End Type
'-------------------------------------
Type CtasAsiento
     DG       As String
     Cta      As String
     TC       As String
     TipoPago As String
     Detalle  As String
     Valor    As Currency
 End Type
'-------------------------------------
 Type Seteos_Documentos
      PosX       As Single
      PosY       As Single
      Tamaño     As Single
      Encabezado As String
 End Type
'-------------------------------------
 Type Formatos_Propios
      Tipo_Objeto As String
      Texto       As String
      color       As String
      Vertical    As Boolean
      Radio       As Single
      Pos_Xo      As Single
      Pos_Yo      As Single
      Pos_Xf      As Single
      Pos_Yf      As Single
      Tamaño      As Single
 End Type
'-------------------------------------
 Type Datos_Giros
      Giro      As String
      Detalle   As String
      Total     As Currency
      TotalAcum As Currency
      Consep    As String
 End Type
'-------------------------------------
 Type DatoMatriz
      Grafico As String
      color As Long
      Texto As String
 End Type
'-------------------------------------
 Type Datos_De_Inventario
      Tipo_SubMod          As String
      Codigo_Inv           As String
      Producto             As String
      Detalle              As String
      Codigo_Barra         As String
      Unidad               As String
      Cta_Inventario       As String
      Cta_Proveedor        As String
      Cta_Costo_Venta      As String
      Cta_Ventas           As String
      Cta_Ventas_0         As String
      Cta_Venta_Anticipada As String
      Patron_Busqueda      As String
      Fecha_Exp            As String
      Fecha_Fab            As String
      Fecha_Stock          As String
      Reg_Sanitario        As String
      Procedencia          As String
      Modelo               As String
      Serie_No             As String
      TC                   As String
      Utilidad             As Single
      Minimo               As Double
      Maximo               As Double
      Stock                As Double
      Costo                As Double
      Valor_Unit           As Double
      PVP                  As Currency
      PVP2                 As Currency
      Div                  As Boolean
      Por_Reservas         As Boolean
      IVA                  As Boolean
      Con_Kardex           As Boolean
 End Type
'-------------------------------------
 Type Progreso_Barras
      PosX_Pict    As Byte
      Puntos       As Byte
      color        As Byte
      Incremento   As Long
      Valor_Maximo As Long
      Mensaje_Box  As String
 End Type
'-------------------------------------
 Type Tipo_Facturas
      T                As String
      TC               As String
      Porc_IVA_S       As String
      Tipo_PRN         As String
      CodigoC          As String
      CodigoB          As String
      CodigoA          As String
      CodigoDr         As String
      Grupo            As String
      Curso            As String
      Cliente          As String
      Contacto         As String
      CI_RUC           As String 'Solo Clientes
      TD               As String
      Razon_Social     As String
      RUC_CI           As String 'Clientes Matriculas
      TB               As String
      DireccionC       As String
      DireccionS       As String
      CiudadC          As String
      DirNumero        As String
      TelefonoC        As String
      EmailC           As String
      EmailC2          As String
      EmailR           As String
      Forma_Pago       As String
      Ejecutivo_Venta  As String
      Cta_CxP          As String
      Cta_CxP_Anterior As String
      Cta_Venta        As String
      Cod_Ejec         As String
      Vendedor         As String
      Afiliado         As String
      Digitador        As String
      Nivel            As String
      Nota             As String
      Observacion      As String
      Definitivo       As String
      Codigo_T         As String
      CodigoU          As String
      Declaracion      As String
      SubCta           As String
      Hora             As String
      Hora_FA          As String
      Hora_NC          As String
      Hora_GR          As String
      Hora_LC          As String
      Serie            As String
      Serie_R          As String
      Serie_NC         As String
      Serie_GR         As String
      Serie_LC         As String
      Autorizacion     As String
      Autorizacion_R   As String
      Autorizacion_NC  As String
      Autorizacion_GR  As String
      Autorizacion_LC  As String
      Fecha_Tours      As String
      ClaveAcceso      As String
      ClaveAcceso_NC   As String
      ClaveAcceso_GR   As String
      ClaveAcceso_LC   As String
      Fecha            As String
      Fecha_C          As String
      Fecha_V          As String
      Fecha_NC         As String
      Fecha_Aut        As String
      Fecha_Aut_NC     As String
      Fecha_Aut_GR     As String
      Fecha_Aut_LC     As String
      Fecha_Corte      As String
      Fecha_Desde      As String
      Fecha_Hasta      As String
      Vencimiento      As String
      FechaGRE         As String
      FechaGRI         As String
      FechaGRF         As String
      CiudadGRI        As String
      CiudadGRF        As String
      Comercial        As String
      CIRUCComercial   As String
      Entrega          As String
      CIRUCEntrega     As String
      Dir_PartidaGR    As String
      Dir_EntregaGR    As String
      Pedido           As String
      Zona             As String
      Placa_Vehiculo   As String
      Error_SRI        As String
      Estado_SRI       As String
      Estado_SRI_NC    As String
      Estado_SRI_GR    As String
      Estado_SRI_LC    As String
      Lugar_Entrega    As String
      DireccionEstab   As String
      NombreEstab      As String
      TelefonoEstab    As String
      LogoTipoEstab    As String
      TP               As String
      Tipo_Pago        As String
      Tipo_Pago_Det    As String
      Tipo_Comp        As String
      Cod_CxC          As String
      CxC_Clientes     As String
      LogoFactura      As String
      LogoNotaCredito  As String
      PDF_ClaveAcceso  As String
      Orden_Compra     As String
      Recibo_No        As String
      
      C                As Boolean
      p                As Boolean
      SP               As Boolean
      ME_              As Boolean
      Com_Pag          As Boolean
      Educativo        As Boolean
      Imp_Mes          As Boolean
      Si_Existe_Doc    As Boolean
      Nuevo_Doc        As Boolean
      EsPorReembolso   As Boolean
      
      Gavetas          As Byte
      
      CantFact         As Integer
      TDT              As Integer
      
      Factura          As Long
      Desde            As Long
      Hasta            As Long
      DAU              As Long
      FUE              As Long
      Remision         As Long
      Solicitud        As Long
      Retencion        As Long
      Nota_Credito     As Long
      Numero           As Long
      
      Porc_C           As Single
      Cotizacion       As Single
      Porc_NC          As Single
      Porc_IVA         As Single
      AltoFactura      As Single
      AnchoFactura     As Single
      EspacioFactura   As Single
      Pos_Factura      As Single
      Pos_Copia        As Single
      
      SubTotal         As Currency
      SubTotal_NC      As Currency
      SubTotal_NCX     As Currency
      Sin_IVA          As Currency
      Con_IVA          As Currency
      Total_Sin_No_IVA As Currency
      Total_Descuento  As Currency
      Total_IVA        As Currency
      Total_IVA_NC     As Currency
      Total_Abonos     As Currency
      Descuento        As Currency
      Descuento2       As Currency
      Descuento_0      As Currency
      Descuento_X      As Currency
      Descuento_NC     As Currency
      Comision         As Currency
      Servicio         As Currency
      Propina          As Currency
      Total_MN         As Currency
      Total_ME         As Currency
      Saldo_MN         As Currency
      Saldo_ME         As Currency
      Cantidad         As Currency
      Kilos            As Currency
      Saldo_Actual     As Currency
      Efectivo         As Currency
      Saldo_Pend       As Currency
      Saldo_Pend_MN    As Currency
      Saldo_Pend_ME    As Currency
      Ret_Fuente       As Currency
      Ret_IVA          As Currency
 End Type
'-------------------------------------
 Type Tipo_Abono
      T               As String
      TP              As String
      Tipo_Cta        As String
      Tipo_Pago       As String
      Cta             As String
      Cta_CxP         As String
      Fecha           As String
      Recibo_No       As String
      Comprobante     As String
      CodigoC         As String
      Recibi_de       As String
      CI_RUC_Cli      As String
      Banco           As String
      Cheque          As String
      Codigo_Inv      As String
      Serie           As String
      Autorizacion    As String
      Establecimiento As String
      Emision         As String
      AutorizacionR   As String
      Tipo_Recibo     As String
      Serie_NC        As String
      Serie_R         As String
      Autorizacion_NC As String
      ME_             As Boolean
      Protestado      As Boolean
      Porcentaje      As Single
      Factura         As Long
      Nota_Credito    As Long
      Secuencial_R    As Long
      Cotizacion      As Double
      Abono           As Currency
 End Type
'-------------------------------------
 Type Tipo_Contribuyente
      Existe             As Boolean
      Estado             As String
      RazonSocial        As String
      RUC_SRI            As String
      NombreComercial    As String
      ClaseRUC           As String
      TipoRUC            As String
      Obligado           As String
      ActividadEconomica As String
      FechaInicio        As String
      FechaCese          As String
      FechaReinicio      As String
      FechaActualización As String
      Categoria          As String
      AgenteRetencion    As String
      MicroEmpresa       As String
 End Type
'-------------------------------------
 Type Tipo_Rol_Pago_Individual
      T               As String
      Tipo_Rubro      As String
      Grupo_Rol       As String
      Dias            As String
      Horas           As String
      Fecha_D         As String
      Fecha_H         As String
      Codigo          As String
      Empleado        As String
      TC              As String
      Cta             As String
      Cheq_Dep_Transf As String
      Codigo_Banco    As String
      SubModulo       As String
      DetSubModulo    As String
      Detalle         As String
      Cod_Rol_Pago    As String
      Retencion_No    As Long
      ID              As Long
      Ingresos        As Currency
      Egresos         As Currency
      Porc            As Single
 End Type
'-------------------------------------
 Type Cyber_Tiempo
    Desde As Byte
    Hasta As Byte
    Valor As Currency
 End Type
'-------------------------------------
 Type Cuentas_Prestamos
      Cta_P_1_30 As String
      Cta_P_31_90 As String
      Cta_P_91_180 As String
      Cta_P_181_360 As String
      Cta_P_Mas_360 As String
      
      Cta_V_1_30 As String
      Cta_V_31_90 As String
      Cta_V_91_180 As String
      Cta_V_181_360 As String
      Cta_V_Mas_360 As String
      
      Cta_Int_Mora As String
      Cta_Seg_Desg_C As String
      Cta_Seg_Desg_P As String
      Cta_Gas_Oper As String
      
      Total_1_30 As Currency
      Total_31_90 As Currency
      Total_91_180 As Currency
      Total_181_360 As Currency
      Total_Mas_360 As Currency
 End Type
 '-------------------------------------
 Type TipoMaterias
      CodigoMat As String
      Materias  As String
      Valor     As Currency
      ValorPQ   As Currency
      ValorSQ   As Currency
      ValorTQ   As Currency
      CantMat   As Byte
 End Type
 '-------------------------------------
Type TipoProyectar
     Proyectado As Currency
     Procesado  As Currency
     Diferencia As Currency
 End Type
 '-------------------------------------
 Type Tipo_Datos_Curso
      Curso              As String
      Curso_Anio         As String
      Curso_Texto        As String
      Curso_Superior     As String
      Paralelo           As String
      Descripcion        As String
      Bachiller          As String
      Especialidad       As String
      Figura_Profesional As String
      Ciclo              As String
      Titulo             As String
      Tipo_Titulo        As String
      Codigo_Titulo      As String
      Seccion            As String
      Nombre_Largo       As String
      CodigoC()          As String
      Alumno()           As String
      Sexo()             As String
      CI_RUC()           As String
      MateriaT()         As String
      CodMatPT()         As String
      Materia()          As String
      CodMat()           As String
      
      ContAlumnos        As Integer
      CantNotas          As Integer
      MateriasPorPagina  As Integer
      
      NotaPQ()           As Currency
      NotaSQ()           As Currency
      NotaTQ()           As Currency
      NotaFinal()        As Currency
      
      PosXMat()          As Single
      
      ContMat            As Byte
      ContMatT           As Byte
 End Type
'-------------------------------------
Type Tipo_Equivalencias
     Desde                    As Single
     Hasta                    As Single
     Letras                   As String
     Cualitativa              As String
     Cualitativa2             As String
     Rango                    As String
     Equivalencia             As String
     Significado_Letras       As String
     Significado_Letras2      As String
     Significado_Evaluacion   As String
     Significado_Evaluacion2  As String
     Significado_Equivalencia As String
End Type
'-------------------------------------
 Type Tipo_De_Identificacion
      RUC_CI             As String
      Codigo_RUC_CI      As String
      Digito_Verificador As String
      Tipo_Beneficiario  As String
      MicroEmpresa       As String
      AgenteRetencion    As String
      RUC_Natural        As Boolean
 End Type
'-------------------------------------
 Type Tipo_Mail
     'Datos para enviar
      para                 As String
      Destinatario         As String
      Asunto               As String
      Adjunto              As String
      Carpeta              As String
      Archivo              As String
      Mensaje              As String
      MensajeHTML          As String
     'Datos del servidor para enviar
      servidor             As String
      ehlo                 As String
      Puerto               As Integer
      useAuntentificacion  As Boolean
      ssl                  As Boolean
      tls                  As Boolean
      Usuario              As String
      Password             As String
      de                   As String
      Remitente            As String
      ListaError           As String
      TipoDeEnvio          As String
     'Posicion de la Lista de mails
      ListaMail            As Byte
      ContadorTiempo       As Byte
      Credito_No           As String
      Volver_Envial        As Boolean
 End Type
'-------------------------------------
 Type Tipo_Lista_Mail
     'Datos para enviar
      Correo_Electronico As String
      Contraseña As String
 End Type
'-------------------------------------
Type Directorio_Dialogo
     Filter      As String
     Title       As String
     InitDir     As String
     FinDir      As String
     File        As String
     Filename    As String
     FilterIndex As Long
End Type
'-------------------------------------
Type Datos_Naciones
     Descripcion As String
     Codigo      As String
     Pais        As String
     Provincia   As String
     CPais       As String
     CRegion     As String
     CProvincia  As String
     CCiudad     As String
     Tipo_Rubro  As String
End Type
'-------------------------------------
Type Concepto_Retencion_ATS
     Codigo              As String
     Concepto            As String
     Porcentaje          As Single
     Ingresar_Porcentaje As String
     Fecha_Inicio        As String
     Fecha_Final         As String
     T                   As String
End Type
'-------------------------------------
Type Tipo_Recibo
     Recibo_No           As String
     Fecha               As String
     SubTotal            As Currency
     IVA                 As Currency
     Total               As Currency
     Saldo               As Currency
     Cobrado_a           As String
     CI_RUC              As String
     Concepto            As String
     Tipo_Recibo         As String
     CodUsuario          As String
End Type
'-------------------------------------

Type Tipo_DBF_Alumnos
     Estado          As String
     Sexo            As String
     Nombres         As String
     Curso           As String
     NombreCurso     As String
     Paralelo        As String
     TB              As String
     bus             As String
     Periodo_Lectivo As String
     Fecha_Nac       As String
     
     codest          As String
     cedula          As String
     cedular         As String
     fonopaga        As String
     pagador         As String
     direcpaga       As String
     emailpaga       As String

     retirado        As Boolean
     pagado          As Boolean
     matriculado     As Boolean
     Aprobado        As Boolean
End Type
'-------------------------------------
Type Tipo_Beneficiarios
     T               As String
     Codigo          As String
     TP              As String
     CI_RUC          As String
     TD              As String
     Fecha           As String
     Fecha_A         As String
     Fecha_N         As String
     Fecha_Cad       As String
     Cliente         As String
     Sexo            As String
     Email1          As String
     Email2          As String
     EmailR          As String
     Direccion       As String
     DirNumero       As String
     Telefono        As String
     Telefono1       As String
     TelefonoT       As String
     Celular         As String
     Ciudad          As String
     Prov            As String
     Pais            As String
     Actividad       As String
     Profesion       As String
     Representante   As String
     RUC_CI_Rep      As String
     TD_Rep          As String
     Direccion_Rep   As String
     Grupo_No        As String
     Contacto        As String
     Calificacion    As String
     Plan_Afiliado   As String
     Cte_Ahr_Otro    As String
     Cta_Transf      As String
     Cta_Numero      As String
     Cta_CxP         As String
     Tipo_Cta        As String
     Cod_Ejec        As String
     Patron_Busqueda As String
     Archivo_Foto    As String
     
     Cod_Banco       As Integer
     Credito         As Integer
     
     Salario         As Currency
     Saldo_Pendiente As Currency
     Total_Anticipo  As Currency
     
     FA              As Boolean
     Asignar_Dr      As Boolean
     Descuento       As Boolean
End Type
'-------------------------------------
Type Tipo_Cta_Descuadre
     Cta          As String
     Fecha        As String
     Total_CxCD   As Currency
     Total_CxCH   As Currency
     Total_Debe   As Currency
     Total_Haber  As Currency
End Type
'-------------------------------------
Type Tipo_Base_DBF
     Entidad       As String
     Tipo_Base     As String
     Carpeta       As String
     Actuales      As String
     Usuario       As String
     Clave         As String
     Curso         As String
     Especialidad  As String
     Paralelo      As String
     Antiguos      As String
     Nuevos        As String
     Periodo       As String
     FechaI        As String
     FechaF        As String
     Cod_Mat_Ini   As String
     Cod_Mat_EBG   As String
     Cod_Mat_Bach  As String
     Cod_Pen_Ini   As String
     Cod_Pen_EBG   As String
     Cod_Pen_Bach  As String
     Val_Mat_Ini   As Currency
     Val_Mat_EBG   As Currency
     Val_Mat_Bach  As Currency
     Val_Pen_Ini   As Currency
     Val_Pen_EBG   As Currency
     Val_Pen_Bach  As Currency
     Puerto        As Long
     Mes_Mat       As Byte
     Base_Datos    As Database
     Registo       As Recordset
     Base_Datos_DB As ADODB.Connection
     Registo_DB    As ADODB.Recordset
End Type
'-------------------------------------
Type Tipo_Conexion
     Entidad       As String
     Tipo_Base     As String
     IP_Server     As String
     Base_Datos    As String
     Usuario       As String
     Clave         As String
     Controlador   As String
     Opcion        As Integer
     Puerto        As Integer
End Type
'-------------------------------------
Type Tipo_Espe_DBF
     CodEspe      As String
     Especialidad As String
End Type
'-------------------------------------
Type Tipo_Estado_SRI
     Clave_De_Acceso    As String
     Autorizacion       As String
     Fecha_Autorizacion As String
     Hora_Autorizacion  As String
     Estado_SRI         As String
     Error_SRI          As String
     Documento_XML      As String
End Type
'-------------------------------------
Type Datos_PC
     InterNet     As Boolean
     Max_IP       As Integer
     Nombre_PC    As String
     Status       As String
     IP_PC        As String
     MAC_PC       As String
     WAN_PC       As String
     Lista_IPs()  As String
End Type
'-------------------------------------
Type Tipo_Impreson
     TipoImpresion     As Byte
     OrientacionPagina As Byte
     PorteLetra        As Byte
     NombreArchivo     As String
     TituloArchivo     As String
     TipoLetra         As String
     PaginaA4          As Boolean
     EsCampoCorto      As Boolean
     VerDocumento      As Boolean
End Type

Type Asiento_SRI
     Documento              As String
     Razon_Social_Emisor    As String
     RUC_Emisor             As String
     RUC_Receptor           As String
     Direccion_Emisor       As String
     Fecha_Emision          As String
     Serie                  As String
     Autorizacion           As String
     Cod_Ret                As String
     Serie_Receptor         As String
     FormaPago              As String
     Cod_Sustento           As String
     Cta_Debito             As String
     Cta_Credito            As String
     Cta_IVA_Gasto          As String
     Cta_Ret_Fuente         As String
     Cta_Ret_IVA_B          As String
     Cta_Ret_IVA_S          As String
     SubModulo              As String
     Codigo_B               As String
     Ambiente               As Byte
     CodPorIva              As Byte
     Cod_Ret_IVA_B          As Byte
     Cod_Ret_IVA_S          As Byte
     Comprobante            As Long
     SubTotal               As Currency
     Total_IVA              As Currency
     Total                  As Currency
     Ret_IVA_B              As Currency
     Ret_IVA_S              As Currency
     Ret_Fuente             As Currency
     Porc_Ret               As Single
     Porc_Ret_IVA_B         As Single
     Porc_Ret_IVA_S         As Single
End Type

Type Datos_PDF
     TipoPDF                As String
     Titulo                 As String
     CodigoBeneficiario     As String
     NombreBeneficiario     As String
     DireccionBeneficiario  As String
     TelefonoBeneficiario   As String
     EmailBeneficiario      As String
     ValorTotal             As Currency
End Type

Type List_Files
     File                   As String
     Ext                    As String
End Type

'Formulario Padre o Principal
'-------------------------------------
 Global MDIFormulario As MDIForm
'-------------------------------------
'Tipos Globales
'-------------------------------------
 Global Ba As Bancos
 Global AXML As Asiento_SRI
 Global TRecibo As Tipo_Recibo
 Global Co As Comprobantes
 Global FA As Tipo_Facturas
 Global NC As Tipo_Facturas
 Global TA As Tipo_Abono
 Global DN As Datos_Naciones
 Global TipoSRI As Tipo_Contribuyente
 Global SRI_Autorizacion As Tipo_Estado_SRI
'-------------------------------------
 Global Heads As EncabezadoReporte
 Global DatInv As Datos_De_Inventario
 Global TBeneficiario As Tipo_Beneficiarios
 Global Tipo_RUC_CI As Tipo_De_Identificacion
'-------------------------------------
 Global Dato_DBF As Tipo_Base_DBF
 Global Estudiante_DBF As Tipo_DBF_Alumnos
 Global DBF_Cursos() As String
 Global DBF_Paralelo() As String
 Global DBF_Especialidad() As Tipo_Espe_DBF
'-------------------------------------
 Global ExisteCtas() As String
'-------------------------------------
 Global TitulosGraf As Grafico
 Global Impresora As Printer
 Global PrinterView As PictureBox
 Global TRol_Pagos As Tipo_Rol_Pago_Individual
 Global TMail As Tipo_Mail
 Global Dir_Dialog As Directorio_Dialogo
 Global CR_ATS As Concepto_Retencion_ATS
 Global Lista_De_Correos(7) As Tipo_Lista_Mail
 Global VectMateria() As TipoMaterias
 Global Dato_Curso As Tipo_Datos_Curso
 Global Progreso_Barra As Progreso_Barras
 Global Equivalencias() As Tipo_Equivalencias
 Global cPrint As cImpresion
 Global IP_PC As Datos_PC
 Global VerPDF As Datos_PDF
'-------------------------------------
 Global AdoRegMySQL As ADODB.Recordset
 'Global ComdMySQL As ADODB.Command
 
 Global BDMySQLEst As ADODB.Connection
 Global RecMySQLEst As ADODB.Recordset
 
 Global tPrint As Tipo_Impreson

'Variable para el UDT que almacena cuatro variables par el Rango de datos de la hoha Excel
 Global Rango As T_Rango

