VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FSuscripcion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORMULARIO DE SUSCRIPCION"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextComision 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   3675
      MaxLength       =   5
      TabIndex        =   8
      Top             =   1785
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoSuscripcion 
      Height          =   330
      Left            =   105
      Top             =   2100
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Suscripcion"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4935
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1050
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4935
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   960
   End
   Begin VB.TextBox TextSector 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   3885
      MaxLength       =   16
      TabIndex        =   6
      Top             =   1050
      Width           =   960
   End
   Begin VB.TextBox TextContrato 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2625
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1050
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBHasta 
      Height          =   330
      Left            =   1365
      TabIndex        =   2
      Top             =   1050
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSMask.MaskEdBox MBDesde 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   1050
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   105
      Top             =   1470
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Aux"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoPrest 
      Height          =   330
      Left            =   105
      Top             =   1785
      Visible         =   0   'False
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Prest"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DCPrest 
      Bindings        =   "FSuscrip1.frx":0000
      DataSource      =   "AdoPrest"
      Height          =   315
      Left            =   105
      TabIndex        =   11
      Top             =   105
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Comision %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3675
      TabIndex        =   7
      Top             =   1470
      Width           =   1170
   End
   Begin VB.Label Label33 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sector"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3885
      TabIndex        =   5
      Top             =   735
      Width           =   960
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Contrato No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2625
      TabIndex        =   3
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   2535
   End
End
Attribute VB_Name = "FSuscripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  FechaTexto = FechaSistema
  MiFecha = FechaSistema
  Factura_No = 0
  MiTiempo = Time
'''' 'Elimina los datos de Suscripciones
''''  sSQL = "DELETE * " _
''''       & "FROM TM_Diario_Caja " _
''''       & "WHERE TP = 'Ven' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "DELETE * " _
''''       & "FROM Clientes " _
''''       & "WHERE Mid(Codigo,1,2) = 'FA' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "SELECT MAX(Desde) As FechaD " _
''''       & "FROM TM_Contratos_Suscrip " _
''''       & "WHERE Contrato_No <> '*' "
''''  SelectAdodc AdoPrest, sSQL
''''  If AdoPrest.Recordset.RecordCount > 0 Then FechaTexto = AdoPrest.Recordset.Fields("FechaD")
''''  sSQL = "SELECT MAX(Factura) As FacturaNo " _
''''       & "FROM TM_Facturas " _
''''       & "WHERE Factura <> 0 "
''''  SelectAdodc AdoPrest, sSQL
''''  If AdoPrest.Recordset.RecordCount > 0 Then Factura_No = AdoPrest.Recordset.Fields("FacturaNo")
''''  sSQL = "SELECT MAX(Fecha) As FechaD " _
''''       & "FROM TM_Diario_Caja " _
''''       & "WHERE TP <> 'Ven' "
''''  SelectAdodc AdoPrest, sSQL
''''  If AdoPrest.Recordset.RecordCount > 0 Then MiFecha = AdoPrest.Recordset.Fields("FechaD")
''''  sSQL = "DELETE * " _
''''       & "FROM Prestamos " _
''''       & "WHERE Mid(Credito_No,3,1) <> '-' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "DELETE * " _
''''       & "FROM Trans_Suscripciones " _
''''       & "WHERE Mid(Contrato_No,3,1) <> '-' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "DELETE * " _
''''       & "FROM Facturas " _
''''       & "WHERE Factura <= " & Factura_No & " " _
''''       & "AND TC = 'FA' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "DELETE * " _
''''       & "FROM Detalle_Factura " _
''''       & "WHERE Factura <= " & Factura_No & " " _
''''       & "AND TC = 'FA' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "DELETE * " _
''''       & "FROM Trans_Abonos " _
''''       & "WHERE Fecha <= #" & BuscarFecha(MiFecha) & "# " _
''''       & "AND TP = 'FA' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "DELETE * " _
''''       & "FROM Trans_Prestamos " _
''''       & "WHERE Item = '" & NumEmpresa & "' " _
''''       & "AND TP = 'SEMA' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "DELETE * " _
''''       & "FROM Trans_Suscripciones " _
''''       & "WHERE Item = '" & NumEmpresa & "' " _
''''       & "AND TP <> '.' "
''''  ConectarAdoExecute sSQL
''''  FSuscripcion.Caption = Format(Time - MiTiempo, "HH:MM:SS")
'''' 'Insertamos datos de Suscripciones
''''  sSQL = "INSERT INTO Clientes " _
''''       & "(T,FA,Codigo,Cliente,Fecha_N,Fecha,TD,CI_RUC,Representante,Profesion,Direccion," _
''''       & "Ciudad,Telefono,Celular,FAX,Sexo,Grupo,Porc_C,FactM,Email,CodigoU,CodigoA,Prov,Pais," _
''''       & "Est_Civil,Actividad,Casilla,Lugar_Trabajo,DirNumero,DireccionT,TelefonoT," _
''''       & "No_Dep,AcumuladoC) " _
''''       & "SELECT 'N' As T,-1 As FA,'FA' & Codigo As Codigo1,Cliente,Fecha_N,Fecha_N," _
''''       & "'O' As TD,RUC_CI,Empresa,Profesion,Direccion,Ciudad,Telefono,Celular,FAX,Sexo," _
''''       & "Grupo,Porc_C,FactM,Email,'" & CodigoUsuario & "' As CU,'.' As CA,'17' As Prv," _
''''       & "'593' As Pa,'S' As S,'.' As Act,E,'.' As LugTrab,'S/N' As DNum,Direccion1,'000000000' As TelfT," _
''''       & "0 As NDep,0 As Acum " _
''''       & "FROM TM_Clientes " _
''''       & "WHERE Codigo <> '.' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "INSERT INTO Prestamos " _
''''       & "(ME,T,Cuenta_No,Fecha,Fecha_C,Numero,Credito_No,Cta,Dia,No_Venc,CodigoU,Item,Atencion," _
''''       & "TP,CodigoE,Tasa,Meses,Interes,Capital,Pagos,Saldo_Pendiente,Plazo,Patrimonio,Encaje) " _
''''       & "SELECT 0 As ME1,TT,'FA' & Codigo_C As Codigos,Desde,Hasta,Factura_No,Contrato_No,Area,Contador," _
''''       & "Contador,'" & CodigoUsuario & "' As CodUsu,'" & NumEmpresa & "' As XItem,'.' As Atenc," _
''''       & "'SEMA' As XTP,'.' As CodE,0 As Tas,52 As XMeses,0 As C1,0 As C2,0 As C3," _
''''       & "0 As C4,0 As C5,0 As C6,0 As C7 " _
''''       & "FROM TM_Contratos_Suscrip " _
''''       & "WHERE TT <> 'A' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "INSERT INTO Detalle_Factura " _
''''       & "(T,TC,CodigoC,Factura,Fecha,Codigo,CodigoL,Producto,Cantidad,Reposicion,Precio," _
''''       & "Total,Total_Desc,Total_IVA,Ruta,Ticket,No_Hab,Cod_Ejec,Porc_C,CodigoU,Item) " _
''''       & "SELECT T,'FA' As TC,'FA' & Codigo_C As CodC,Factura_No,Fecha,Codigo,CodigoL,Producto,Cantidad," _
''''       & "Reposicion,Precio,Total,Total_Desc,Total_IVA,'.' As Ruta,'.' As Ticket,'.' As No_Hab,Cod_Ejec," _
''''       & "0 As Porc,'" & CodigoUsuario & "' As CodUsu,'" & NumEmpresa & "' As Item " _
''''       & "FROM TM_Detalle_Factura " _
''''       & "WHERE T <> 'A' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "INSERT INTO Facturas " _
''''       & "(C,ME,T,TC,Factura,CodigoC,Fecha,Fecha_C,Fecha_V,SubTotal,Con_IVA,Sin_IVA,IVA," _
''''       & "Descuento,Porc_C,Comision,Servicio,Total_MN,Total_ME,Saldo_MN,Saldo_ME,Forma_Pago," _
''''       & "Cotizacion,Cta_CxP,Cta_Venta,Cod_Ejec,Nivel,Nota,Observacion,Definitivo,Codigo_T," _
''''       & "Fecha_Tours,CodigoU,Item) " _
''''       & "SELECT 0 As C,ME,T,'FA' As TC,Factura,'FA' & Codigo_C As CodigoC,Fecha,Fecha_C,Fecha_V," _
''''       & "SubTotal,Con_IVA,Sin_IVA,IVA,Descuento,Porc_C,Comision,Servicio,Total_MN,Total_ME," _
''''       & "Saldo_MN,Saldo_ME,Forma_Pago,Cotizacion,Cta_CxC,Cta_Venta,Cod_Ejec,Nivel,'.' As XNota," _
''''       & "'.' As XObservacion,'.' As XDefinitivo,'.' As XCodigo_T,'.' As XFecha_Tours," _
''''       & "'" & CodigoUsuario & "' As CodUsu,'" & NumEmpresa & "' As Item " _
''''       & "FROM TM_Facturas " _
''''       & "WHERE T <> 'A' "
''''  ConectarAdoExecute sSQL
''''  sSQL = "INSERT INTO Trans_Abonos " _
''''       & "(C,ME,T,TP,Cta,Cta_CxP,Fecha,Recibo_No,Comprobante,Factura,Abono,CodigoC,Cotizacion," _
''''       & "Banco,Cheque,CodigoU,Item) " _
''''       & "SELECT C,0 As ME,T,'FA' As TP,'" & Cta_CajaG & "' As Cta1,CtaxCob,Fecha,Diario_No," _
''''       & "Diario_No,Factura,Caja_MN,'FA' & Codigo_C As CodigoC,0 As Cotiz,UCASE(Banco),Cheque," _
''''       & "'" & CodigoUsuario & "' As CodUsu,'" & NumEmpresa & "' As Item " _
''''       & "FROM TM_Diario_Caja " _
''''       & "WHERE Caja_MN <> 0 "
''''  ConectarAdoExecute sSQL
''''  sSQL = "INSERT INTO Trans_Abonos " _
''''       & "(C,ME,T,TP,Cta,Cta_CxP,Fecha,Recibo_No,Comprobante,Factura,Abono,CodigoC,Cotizacion," _
''''       & "Banco,Cheque,CodigoU,Item) " _
''''       & "SELECT C,0 As ME,T,'FA' As TP,'" & Cta_CajaBA & "' As Cta1,CtaxCob,Fecha,Diario_No," _
''''       & "Diario_No,Factura,Caja_Vaucher,'FA' & Codigo_C As CodigoC,0 As Cotiz,UCASE(Banco),Cheque," _
''''       & "'" & CodigoUsuario & "' As CodUsu,'" & NumEmpresa & "' As Item " _
''''       & "FROM TM_Diario_Caja " _
''''       & "WHERE Caja_Vaucher <> 0 "
''''  ConectarAdoExecute sSQL
''''  FSuscripcion.Caption = Format(Time - MiTiempo, "HH:MM:SS")
'''' 'Actualizamos Suscripciones
''''  sSQL = "UPDATE Facturas " _
''''       & "SET C = -1 " _
''''       & "WHERE Item = '" & NumEmpresa & "' " _
''''       & "AND Fecha < #" & BuscarFecha("01/11/2003") & "# "
''''  ConectarAdoExecute sSQL
''''  sSQL = "UPDATE Detalle_Factura " _
''''       & "SET C = -1 " _
''''       & "WHERE Item = '" & NumEmpresa & "' " _
''''       & "AND Fecha < #" & BuscarFecha("01/11/2003") & "# "
''''  ConectarAdoExecute sSQL
''''  sSQL = "UPDATE Trans_Abonos " _
''''       & "SET C = -1 " _
''''       & "WHERE Item = '" & NumEmpresa & "' " _
''''       & "AND Fecha < #" & BuscarFecha("01/11/2003") & "# "
''''  ConectarAdoExecute sSQL
''''  sSQL = "UPDATE Trans_Suscripciones " _
''''       & "SET AC = -1 " _
''''       & "WHERE Item = '" & NumEmpresa & "' " _
''''       & "AND T = 'A' "
''''  ConectarAdoExecute sSQL
  FSuscripcion.Caption = Format(Time - MiTiempo, "HH:MM:SS")
 'Reprocesamos la Suscripcion
  Saldo = 0: Diferencia = 0: Total = 0
  sSQL = "SELECT * " _
       & "FROM Prestamos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY Tasa DESC  "
  SelectDBCombo DCPrest, AdoPrest, sSQL, "Credito_No"
  With AdoPrest.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Total = 0: Contador = 0
       Do While Not .EOF
          Contador = Contador + 1
          FSuscripcion.Caption = Format(Time - MiTiempo, "HH:MM:SS") & ": " & Format(Contador / .RecordCount, "00%")
          T = .Fields("T")
          TipoDoc = .Fields("TP")
          Codigo1 = .Fields("Cta")
          Codigo = .Fields("Cuenta_No")
          Credito_No = .Fields("Credito_No")
          Opcion = .Fields("Dia")
          MiFecha = .Fields("Fecha")
          I = CFechaLong(.Fields("Fecha"))
          J = CFechaLong(.Fields("Fecha_C"))
          If I >= J Then J = I
          Select Case TipoDoc
            Case "MENS": NoMeses = (J - I) / 31
            Case "QUIN": NoMeses = (J - I) / 15
            Case "SEMA": NoMeses = (J - I) / 7
          End Select
          'If NoMeses > 254 Then NoMeses = 254
          If NoMeses <= 0 Then NoMeses = 1
          If Opcion > NoMeses Then Opcion = NoMeses
          Saldo = Round(Total / NoMeses, 2)
          For Cuota_No = 1 To NoMeses
              Select Case TipoDoc
                Case "MENS": MiFecha = SiguienteMes(MiFecha)
                Case "QUIN": MiFecha = CLongFecha(CFechaLong(MiFecha) + 15)
                Case "SEMA": MiFecha = CLongFecha(CFechaLong(MiFecha) + 7)
              End Select
              If Cuota_No <= Opcion Then NoCert = adTrue Else NoCert = adFalse
              If T = "A" Then NoCert = adTrue
              sSQL = "INSERT INTO Trans_Suscripciones " _
                   & "(T,AC,E,Fecha,Fecha_E,TP,Contrato_No,Codigo,Ent_No,Valor_Ed,Comprobante,CodigoU,Item) " _
                   & "VALUES " _
                   & "('" & T & "',0," & NoCert & ",#" & BuscarFecha(MiFecha) & "#," _
                   & "#" & BuscarFecha(MiFecha) & "#,'" & TipoDoc & "','" & Credito_No & "'," _
                   & "'" & Codigo & "'," & Cuota_No & ",0,'" & Val(Credito_No) & "'," _
                   & "'" & CodigoUsuario & "','" & NumEmpresa & "') "
              ConectarAdoExecute sSQL
          Next Cuota_No
         .MoveNext
       Loop
   End If
  End With
  MsgBox "Fin del Proceso"
  Unload FSuscripcion
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
  Trans_No = 249
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FSuscripcion
  ConectarAdodc AdoAux
  ConectarAdodc AdoPrest
  ConectarAdodc AdoSuscripcion
End Sub

Private Sub MBDesde_GotFocus()
  MarcarTexto MBDesde
End Sub

Private Sub MBDesde_LostFocus()
  FechaValida MBDesde
End Sub

Private Sub MBHasta_GotFocus()
  MarcarTexto MBHasta
End Sub

Private Sub MBHasta_LostFocus()
  FechaValida MBHasta
End Sub

Private Sub TextContrato_GotFocus()
  MarcarTexto TextContrato
End Sub

Private Sub TextContrato_LostFocus()
  TextoValido TextContrato, , True
  TextContrato.Text = Format(TextContrato.Text, "0000000")
  Credito_No = TextContrato.Text
End Sub

Private Sub TextSector_GotFocus()
  MarcarTexto TextSector
End Sub

Private Sub TextSector_LostFocus()
  TextoValido TextSector, , True
End Sub

