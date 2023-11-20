VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FCambiosCodigos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   2850
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPorc 
      Height          =   330
      Left            =   5355
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   1680
      Width           =   1170
   End
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "CambiCod.frx":0000
      DataSource      =   "AdoFactura"
      Height          =   315
      Left            =   4725
      TabIndex        =   7
      Top             =   525
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Factura"
   End
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "CambiCod.frx":0019
      DataSource      =   "AdoCliente"
      Height          =   315
      Left            =   105
      TabIndex        =   14
      Top             =   2415
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Clientes"
   End
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   105
      Top             =   2835
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   3885
      Top             =   2835
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Factura"
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
      Caption         =   "&Salir"
      Height          =   855
      Left            =   6615
      Picture         =   "CambiCod.frx":0032
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1050
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cambiar"
      Height          =   855
      Left            =   6615
      Picture         =   "CambiCod.frx":08FC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   105
      Width           =   960
   End
   Begin MSDataListLib.DataCombo DCTipo 
      Bindings        =   "CambiCod.frx":0D3E
      DataSource      =   "AdoDetAcomp"
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   525
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "FA/NV"
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
   Begin MSDataListLib.DataCombo DCSerie 
      Bindings        =   "CambiCod.frx":0D58
      DataSource      =   "AdoSerie"
      Height          =   315
      Left            =   735
      TabIndex        =   3
      Top             =   525
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "001001"
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
   Begin MSDataListLib.DataCombo DCAutorizacion 
      Bindings        =   "CambiCod.frx":0D6F
      DataSource      =   "AdoAutorizacion"
      Height          =   360
      Left            =   1365
      TabIndex        =   1
      Top             =   105
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "1234567890"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   5775
      Top             =   2835
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "DetAcomp"
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
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   1995
      Top             =   2835
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Serie"
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
   Begin MSAdodcLib.Adodc AdoAutorizacion 
      Height          =   330
      Left            =   6615
      Top             =   2415
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
      Caption         =   "Autorizacion"
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Porcentaje"
      Height          =   330
      Left            =   5355
      TabIndex        =   11
      Top             =   1365
      Width           =   1170
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Tipo "
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1995
      TabIndex        =   4
      Top             =   525
      Width           =   540
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " S&erie"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   525
      Width           =   645
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Autorización"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label LblObs 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
      ForeColor       =   &H00C000C0&
      Height          =   645
      Left            =   105
      TabIndex        =   10
      Top             =   1365
      Width           =   5160
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   4725
      TabIndex        =   9
      Top             =   945
      Width           =   1800
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Factura No."
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3465
      TabIndex        =   6
      Top             =   525
      Width           =   1275
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   945
      Width           =   4635
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CAMBIAR ESTE DOCUMENTO POR EL SIGUIENTE &BENEFICIARIO"
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   2100
      Width           =   6420
   End
End
Attribute VB_Name = "FCambiosCodigos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Select Case Opciones
    Case 1, 2
        With AdoCliente.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
            .Find ("Cliente Like '" & DCCliente.Text & "' ")
             If Not .EOF Then
                CodigoCliente = .Fields("Codigo")
                NombreCliente = DCCliente.Text
                Mensajes = "Cambiar: " & LblObs.Caption & " (" & CodigoP & ")" & vbCrLf & vbCrLf _
                         & "Por: " & NombreCliente & " (" & CodigoCliente & ")" & vbCrLf & vbCrLf _
                         & "De la Autorizacion siguiente: " & Autorizacion & " - " & TipoFactura & vbCrLf & vbCrLf _
                         & "Del Documento No. " & SerieFactura & "-" & Factura_No & vbCrLf
                Titulo = "Formulario de Grabación."
                If BoxMensaje = vbYes Then
                  'MsgBox Opciones
                   If Opciones = 1 Then
                      sSQL = "UPDATE Facturas " _
                           & "SET CodigoC = '" & CodigoCliente & "' " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND Factura = " & Factura_No & " " _
                           & "AND CodigoC = '" & CodigoP & "' " _
                           & "AND Autorizacion = '" & Autorizacion & "' " _
                           & "AND Serie = '" & SerieFactura & "' " _
                           & "AND TC = '" & TipoFactura & "' "
                      Ejecutar_SQL_SP sSQL
                      sSQL = "UPDATE Detalle_Factura " _
                           & "SET CodigoC = '" & CodigoCliente & "' " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND Factura = " & Factura_No & " " _
                           & "AND CodigoC = '" & CodigoP & "' " _
                           & "AND Autorizacion = '" & Autorizacion & "' " _
                           & "AND Serie = '" & SerieFactura & "' " _
                           & "AND TC = '" & TipoFactura & "' "
                      Ejecutar_SQL_SP sSQL
                      sSQL = "UPDATE Trans_Abonos " _
                           & "SET CodigoC = '" & CodigoCliente & "' " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND Factura = " & Factura_No & " " _
                           & "AND CodigoC = '" & CodigoP & "' " _
                           & "AND Autorizacion = '" & Autorizacion & "' " _
                           & "AND Serie = '" & SerieFactura & "' " _
                           & "AND TP = '" & TipoFactura & "' "
                      Ejecutar_SQL_SP sSQL
                   Else
                      sSQL = "UPDATE Facturas " _
                           & "SET Cod_Ejec = '" & CodigoCliente & "',Porc_C = " & Redondear(Val(TxtPorc) / 100, 4) & " " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND Factura = " & Factura_No & " " _
                           & "AND Cod_Ejec = '" & CodigoP & "' " _
                           & "AND Autorizacion = '" & Autorizacion & "' " _
                           & "AND Serie = '" & SerieFactura & "' " _
                           & "AND TC = '" & TipoFactura & "' "
                      Ejecutar_SQL_SP sSQL
                      sSQL = "UPDATE Detalle_Factura " _
                           & "SET Cod_Ejec = '" & CodigoCliente & "' " _
                           & "WHERE Item = '" & NumEmpresa & "' " _
                           & "AND Periodo = '" & Periodo_Contable & "' " _
                           & "AND Factura = " & Factura_No & " " _
                           & "AND Cod_Ejec = '" & CodigoP & "' " _
                           & "AND Autorizacion = '" & Autorizacion & "' " _
                           & "AND Serie = '" & SerieFactura & "' " _
                           & "AND TC = '" & TipoFactura & "' "
                      Ejecutar_SQL_SP sSQL
                   End If
                End If
             Else
                MsgBox "Cliente no Asignado"
             End If
         Else
             MsgBox "No existen datos"
         End If
        End With
    Case 3
        Numero = Val(DCCliente.Text)
        Mensajes = "Cambiar el numero de la Factura" & vbCrLf & vbCrLf _
                 & "De: " & LblObs.Caption & vbCrLf & vbCrLf _
                 & "Por el Numero: " & Format$(Numero, "0000000") & vbCrLf & vbCrLf _
                 & "De la Autorizacion: " & Autorizacion & " - " & TipoFactura & vbCrLf & vbCrLf _
                 & "Del Documento No. " & SerieFactura & "-" & Factura_No & vbCrLf
        Titulo = "Formulario de Grabación."
        If BoxMensaje = vbYes Then
           TBeneficiario = Leer_Datos_Clientes(CodigoP)
           sSQL = "UPDATE Facturas " _
                & "SET Factura = " & Numero & ", " _
                & "Clave_Acceso = '.', "
           If Len(TBeneficiario.Representante) > 1 Then
              sSQL = sSQL _
                   & "RUC_CI = '" & TBeneficiario.RUC_CI_Rep & "', " _
                   & "TB = '" & TBeneficiario.TD_Rep & "', " _
                   & "Razon_Social = '" & TBeneficiario.Representante & "', "
           End If
           sSQL = sSQL _
                & "Estado_SRI = 'CN' " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Factura = " & Factura_No & " " _
                & "AND CodigoC = '" & CodigoP & "' " _
                & "AND Autorizacion = '" & Autorizacion & "' " _
                & "AND Serie = '" & SerieFactura & "' " _
                & "AND TC = '" & TipoFactura & "' "
           Ejecutar_SQL_SP sSQL
           
           sSQL = "UPDATE Detalle_Factura " _
                & "SET Factura = " & Numero & " " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Factura = " & Factura_No & " " _
                & "AND CodigoC = '" & CodigoP & "' " _
                & "AND Autorizacion = '" & Autorizacion & "' " _
                & "AND Serie = '" & SerieFactura & "' " _
                & "AND TC = '" & TipoFactura & "' "
           Ejecutar_SQL_SP sSQL
           
           sSQL = "UPDATE Trans_Abonos " _
                & "SET Factura = " & Numero & " " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND Periodo = '" & Periodo_Contable & "' " _
                & "AND Factura = " & Factura_No & " " _
                & "AND CodigoC = '" & CodigoP & "' " _
                & "AND Autorizacion = '" & Autorizacion & "' " _
                & "AND Serie = '" & SerieFactura & "' " _
                & "AND TP = '" & TipoFactura & "' "
           Ejecutar_SQL_SP sSQL
           
          'Anulamos la misma factura anterior
           SetAdoAddNew "Facturas"
           SetAdoFields "T", Anulado
           SetAdoFields "TC", TipoFactura
           SetAdoFields "Factura", Factura_No
           SetAdoFields "CodigoC", CodigoP
           SetAdoFields "Cod_CxC", TipoDoc
           SetAdoFields "Cta_CxP", Cta_Cobrar
           SetAdoFields "Serie", SerieFactura
           SetAdoFields "Autorizacion", Autorizacion
           SetAdoFields "Item", NumEmpresa
           SetAdoUpdate
        End If
  End Select
  RatonNormal
End Sub

Private Sub Command2_Click()
   Unload FCambiosCodigos
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCFactura_GotFocus()
  TotalCajaMN = 0
  TotalCajaME = 0
  Total_Bancos = 0
  Total_Tarjeta = 0
  Total_IVA = 0
  Total_Ret = 0
End Sub

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then DCCliente.SetFocus
End Sub

Private Sub DCFactura_LostFocus()
  Factura_No = 0: Saldo = 0: Cotizacion = 0: TotalDolar = 0
  CodigoP = Ninguno:  Saldo_ME = 0
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura = " & Val(DCFactura.Text) & " ")
       If Not .EOF Then
          Factura_No = .Fields("Factura")
          LblObs.Caption = .Fields("Cliente")
          Select Case Opciones
            Case 1, 3: CodigoP = .Fields("CodigoC")
            Case 2: CodigoP = .Fields("Cod_Ejec")
          End Select
          TipoDoc = .Fields("Cod_CxC")
          Cta_Cobrar = .Fields("Cta_CxP")
          Saldo = Redondear(.Fields("Saldo_MN"), 2)
          Saldo_ME = Redondear(.Fields("Saldo_ME"), 2)
          TipoDoc = .Fields("TC")
          TotalDolar = .Fields("Total_ME")
          Cotizacion = .Fields("Cotizacion")
          Efectivo = 0: Cheque = 0: Retencion = 0
          LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
       End If
    End If
  End With
  SaldoDisp = Saldo - TotalCajaMN - TotalCajaME - Total_Bancos - Total_Tarjeta - Total_IVA - Total_Ret
End Sub

Private Sub Form_Activate()
  Mifecha = BuscarFecha(FechaSistema)
  If CodigoCliente = "" Then CodigoCliente = Ninguno
  If Factura_No <= 0 Then Factura_No = 0
 'Listamos las autorizaciones de facturas pendientes por cliente
 '& "AND NOT TC IN ('OP','LC','CP') "
  sSQL = "SELECT Autorizacion " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Autorizacion " _
       & "ORDER BY Autorizacion "
  SelectDB_Combo DCAutorizacion, AdoAutorizacion, sSQL, "Autorizacion"
  Listar_Clientes_F
  RatonNormal
End Sub

Private Sub Form_Load()
   CentrarForm FCambiosCodigos
   ConectarAdodc AdoCliente
   ConectarAdodc AdoSerie
   ConectarAdodc AdoFactura
   ConectarAdodc AdoAutorizacion
   ConectarAdodc AdoDetAcomp
End Sub

Public Sub Listar_Facturas_P()
  Select Case Opciones
    Case 1, 3
         SQL1 = "SELECT F.*,C.Cliente " _
              & "FROM Facturas As F,Clientes As C " _
              & "WHERE F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "AND F.Autorizacion = '" & Autorizacion & "' " _
              & "AND F.Serie = '" & SerieFactura & "' " _
              & "AND F.TC = '" & TipoFactura & "' " _
              & "AND F.CodigoC = C.Codigo " _
              & "ORDER BY F.Factura "
    Case 2
         SQL1 = "SELECT F.*,C.Cliente " _
              & "FROM Facturas As F,Clientes As C " _
              & "WHERE F.Item = '" & NumEmpresa & "' " _
              & "AND F.Periodo = '" & Periodo_Contable & "' " _
              & "AND F.Autorizacion = '" & Autorizacion & "' " _
              & "AND F.Serie = '" & SerieFactura & "' " _
              & "AND F.TC = '" & TipoFactura & "' " _
              & "AND F.Cod_Ejec = C.Codigo " _
              & "ORDER BY F.Factura "
  End Select
  SelectDB_Combo DCFactura, AdoFactura, SQL1, "Factura", True
End Sub

Public Sub Listar_Clientes_F()
 Select Case Opciones
   Case 1
        sSQL = "SELECT Grupo,Cliente,Codigo,CI_RUC,Direccion,Telefono " _
             & "FROM Clientes " _
             & "WHERE Cliente <> '.' " _
             & "ORDER BY Cliente "
        SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
   Case 2
        sSQL = "SELECT C.Grupo,C.Cliente,C.Codigo,C.CI_RUC,C.Direccion,C.Telefono,CRP.Salario " _
             & "FROM Clientes As C,Catalogo_Rol_Pagos As CRP " _
             & "WHERE CRP.Item = '" & NumEmpresa & "' " _
             & "AND CRP.Periodo = '" & Periodo_Contable & "' " _
             & "AND C.Codigo = CRP.Codigo " _
             & "ORDER BY C.Cliente "
        SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
   Case 3
        sSQL = "SELECT Factura " _
             & "FROM Facturas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "AND T = 'A' " _
             & "GROUP BY Factura " _
             & "ORDER BY Factura DESC "
        SelectDB_Combo DCCliente, AdoCliente, sSQL, "Factura"
 End Select
End Sub

Private Sub DCAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCAutorizacion_LostFocus()
 'Listamos las autorizaciones de facturas pendientes por cliente
  Autorizacion = DCAutorizacion
  If Autorizacion = "" Then Autorizacion = Ninguno
  sSQL = "SELECT Serie " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND NOT TC IN ('OP','C','P') " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & "GROUP BY Serie " _
       & "ORDER BY Serie "
  SelectDB_Combo DCSerie, AdoSerie, sSQL, "Serie"
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
  SerieFactura = DCSerie
  If SerieFactura = "" Then SerieFactura = Ninguno
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND NOT TC IN ('OP','C','P') " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "GROUP BY TC " _
       & "ORDER BY TC "
  SelectDB_Combo DCTipo, AdoDetAcomp, sSQL, "TC"
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
  TipoFactura = DCTipo
  If TipoFactura = "" Then TipoFactura = Ninguno
  Listar_Facturas_P
End Sub

Private Sub TxtPorc_GotFocus()
  MarcarTexto TxtPorc
End Sub

Private Sub TxtPorc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

