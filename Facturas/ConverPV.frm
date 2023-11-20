VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FConvertirPV 
   Caption         =   "INGRESO DE CAJA"
   ClientHeight    =   7155
   ClientLeft      =   5040
   ClientTop       =   4395
   ClientWidth     =   11355
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DGFactura 
      Bindings        =   "ConverPV.frx":0000
      Height          =   4950
      Left            =   1155
      TabIndex        =   15
      Top             =   1785
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8731
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PictBarra 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   105
      ScaleHeight     =   585
      ScaleWidth      =   1950
      TabIndex        =   20
      Top             =   4725
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Imprimir"
      Height          =   750
      Left            =   105
      Picture         =   "ConverPV.frx":0019
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2205
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Consultar"
      Height          =   750
      Left            =   105
      Picture         =   "ConverPV.frx":08E3
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   525
      Width           =   960
   End
   Begin VB.OptionButton OpcIndividual 
      Caption         =   "Individual Ticket No."
      Height          =   330
      Left            =   6195
      TabIndex        =   4
      Top             =   105
      Width           =   2115
   End
   Begin VB.OptionButton OpcResumen 
      Caption         =   "Resumen"
      Height          =   330
      Left            =   4830
      TabIndex        =   3
      Top             =   105
      Value           =   -1  'True
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   1785
      Top             =   2625
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
      Caption         =   "IngCaja"
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
      Left            =   1155
      Top             =   6720
      Width           =   6735
      _ExtentX        =   11880
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
      Height          =   750
      Left            =   105
      Picture         =   "ConverPV.frx":0D25
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3045
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   750
      Left            =   105
      Picture         =   "ConverPV.frx":15EF
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1365
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   2100
      TabIndex        =   1
      Top             =   105
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   1785
      Top             =   2940
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
      Caption         =   "Banco"
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
   Begin MSDataListLib.DataCombo DCBanco 
      Bindings        =   "ConverPV.frx":1A31
      DataSource      =   "AdoBanco"
      Height          =   315
      Left            =   4620
      TabIndex        =   9
      Top             =   945
      Visible         =   0   'False
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Banco"
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   3360
      TabIndex        =   2
      Top             =   105
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
   Begin MSDataListLib.DataCombo DCTicket 
      Bindings        =   "ConverPV.frx":1A48
      DataSource      =   "AdoTicket"
      Height          =   315
      Left            =   8505
      TabIndex        =   5
      Top             =   105
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Banco"
   End
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   1785
      Top             =   2310
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
      Caption         =   "Linea"
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
   Begin MSDataListLib.DataCombo DCLinea 
      Bindings        =   "ConverPV.frx":1A60
      DataSource      =   "AdoLinea"
      Height          =   315
      Left            =   4620
      TabIndex        =   7
      Top             =   525
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CxC Clientes"
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
   Begin MSAdodcLib.Adodc AdoTicket 
      Height          =   330
      Left            =   1785
      Top             =   3255
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
      Caption         =   "Ticket"
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "ConverPV.frx":1A77
      DataSource      =   "AdoCliente"
      Height          =   315
      Left            =   4620
      TabIndex        =   11
      Top             =   1365
      Visible         =   0   'False
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Banco"
   End
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   1785
      Top             =   3570
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
      Caption         =   "Ticket"
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
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "001-001-0000001"
      Height          =   330
      Left            =   9240
      TabIndex        =   18
      Top             =   525
      Width           =   2010
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cambio de Nombre del Cliente"
      Height          =   330
      Left            =   1155
      TabIndex        =   10
      Top             =   1365
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Seleccione el Tipo de &Facturación"
      Height          =   330
      Left            =   1155
      TabIndex        =   6
      Top             =   525
      Width           =   3480
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Seleccion donde &Ingresa lo recaudado"
      Height          =   330
      Left            =   1155
      TabIndex        =   8
      Top             =   945
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   9450
      TabIndex        =   16
      Top             =   6720
      Width           =   1800
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
      Height          =   330
      Left            =   7875
      TabIndex        =   17
      Top             =   6720
      Width           =   1590
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA DE CIERRE:"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2010
   End
End
Attribute VB_Name = "FConvertirPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CierreDelDia()
  sSQL = "SELECT Fecha,Ticket " _
       & "FROM Trans_Ticket " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
       & "AND C = " & Val(adFalse) & " " _
       & "ORDER BY Fecha,Ticket "
  Select_Adodc AdoIngCaja, sSQL
  With AdoIngCaja.Recordset
   If .RecordCount > 0 Then
       MsgBox "Cierre del día: " & .Fields("Fecha") & "(" & .Fields("Ticket") & ")" & vbCrLf
       MBFechaF = .Fields("Fecha")
       MBFechaI = .Fields("Fecha")
   End If
  End With
  MBFechaI.SetFocus
End Sub

Private Sub Command1_Click()
  FechaValida MBFechaI
  FechaValida MBFechaF
  If OpcIndividual.value Then Cta_Aux = SinEspaciosIzq(DCBanco) Else Cta_Aux = Cta_CajaG
  Numero = Val(DCTicket.Text)
  FA.Cod_CxC = DCLinea
  Lineas_De_CxC FA
  TA.Serie = FA.Serie
  TA.Autorizacion = FA.Autorizacion
  FA.Nuevo_Doc = True
  FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
  Mensajes = "Esta Seguro que desea convertir Tickes por " & FA.TC & ", No. " & FA.Factura
  Titulo = "Formulario de Grabación."
  If BoxMensaje = vbYes Then
     'Control_Procesos  Normal, "Conversion de PV desde " & Factura_Desde & " hasta " & Factura_Hasta
      Encerar_Factura FA
      sSQL = "DELETE * " _
           & "FROM Asiento_F " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND CodigoU = '" & CodigoUsuario & "' "
     Ejecutar_SQL_SP sSQL
     FA.Cod_CxC = DCLinea
     Lineas_De_CxC FA
     FechaTexto = MBFecha ' FechaSistema
     DGFactura.Visible = False
     Total = 0: Contador = 0: Ln_No = 0
     With AdoFactura.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
          If OpcResumen.value Then FA.CodigoC = .Fields("CodigoC") Else FA.CodigoC = CodigoCli
          Mifecha = .Fields("Fecha")
          FechaInicial = .Fields("Fecha")
          FechaFinal = .Fields("Fecha")
          FA.T = Cancelado
          FA.Fecha = .Fields("Fecha")
          FA.Fecha_C = FA.Fecha
          FA.Fecha_V = FA.Fecha
          Do While Not .EOF
            'MsgBox Ln_No & vbCrLf & Cant_Item_FA
             If (Ln_No >= Cant_Item_FA) Or (Mifecha <> .Fields("Fecha")) Or (FA.CodigoC <> .Fields("CodigoC")) Then
               'Empezamos a grabar la factura
                Calculos_Totales_Factura FA
                Cta_Cobrar = FA.Cta_CxP
                FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, True)
                FA.T = "C": FA.Saldo_MN = 0
                'MsgBox FA.Factura
                Grabar_Factura FA, True
               'Forma del Abono
                TA.T = Normal
                TA.TP = FA.TC
                TA.Fecha = FA.Fecha
                TA.CodigoC = FA.CodigoC
                TA.Cta = Cta_Aux
                TA.Banco = "EFECTIVO MN"
                TA.Cheque = "TICKET"
                TA.Factura = FA.Factura
                TA.Abono = FA.Total_MN
                Grabar_Abonos TA
               'Limpiamos para otra factura/Nota de Venta
                Encerar_Factura FA
                sSQL = "DELETE * " _
                     & "FROM Asiento_F " _
                     & "WHERE Item = '" & NumEmpresa & "' " _
                     & "AND CodigoU = '" & CodigoUsuario & "' "
                Ejecutar_SQL_SP sSQL
                FA.Cod_CxC = DCLinea
                Lineas_De_CxC FA
                TA.Serie = FA.Serie
                TA.Autorizacion = FA.Autorizacion
                FA.T = Cancelado
                FA.Fecha = .Fields("Fecha")
                FA.Fecha_C = FA.Fecha
                FA.Fecha_V = FA.Fecha
                Ln_No = 0
                If OpcResumen.value Then FA.CodigoC = .Fields("CodigoC") Else FA.CodigoC = CodigoCli
                Mifecha = .Fields("Fecha")
                If Cant_Item_FA <= 0 Then Cant_Item_FA = 15
             End If
             Contador = Contador + 1
             FConvertirPV.Caption = Format$(Contador / .RecordCount, "00%")
            'Factura_No = .Fields("Factura")
             Codigo_Inv = .Fields("Codigo_Inv")
             Cantidad = .Fields("Cant")
             Precio = .Fields("PVP")
             Total = .Fields("Total_PV")
             FechaFinal = .Fields("Fecha")
            'Gabamos el detalle de la factura
             SetAdoAddNew "Asiento_F"
             SetAdoFields "CODIGO", Codigo_Inv
             SetAdoFields "CODIGO_L", FA.Cod_CxC
             SetAdoFields "PRODUCTO", Ninguno
             SetAdoFields "CANT", Cantidad
             SetAdoFields "PRECIO", Precio
             SetAdoFields "TOTAL", Total
             SetAdoFields "Total_IVA", Real3
             SetAdoFields "Cta", FA.Cta_Venta
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoFields "A_No", CByte(Ln_No)
             SetAdoUpdate
             Ln_No = Ln_No + 1
            .MoveNext
          Loop
      End If
     End With
     
    'Empezamos a grabar la factura
     Calculos_Totales_Factura FA
     Cta_Cobrar = FA.Cta_CxP
     FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, True)
     FA.T = "C": FA.Saldo_MN = 0
     Grabar_Factura FA, True
     TA.T = Normal
     TA.TP = FA.TC
     TA.Fecha = FA.Fecha
     TA.CodigoC = FA.CodigoC
     TA.Cta = Cta_Aux
     TA.Banco = "EFECTIVO MN"
     TA.Cheque = "TICKET"
     TA.Factura = FA.Factura
     TA.Abono = FA.Total_MN
     Grabar_Abonos TA
    'Desactivamos los puntos de venta pendientes seleccionados
     FechaIni = BuscarFecha(FechaInicial)
     FechaFin = BuscarFecha(FechaFinal)
     sSQL = "UPDATE Trans_Ticket " _
          & "SET C = " & Val(adTrue) & " " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND C = " & Val(adFalse) & " " _
          & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
     If OpcIndividual.value Then sSQL = sSQL & "AND Ticket = " & Numero & " "
     Ejecutar_SQL_SP sSQL
    'Actualizamos el detalle de las Ventas del día
     If SQL_Server Then
        sSQL = "UPDATE Detalle_Factura " _
             & "SET Producto = CP.Producto " _
             & "FROM Detalle_Factura As DF,Catalogo_Productos As CP "
     Else
        sSQL = "UPDATE Detalle_Factura As DF,Catalogo_Productos As CP " _
             & "SET DF.Producto = CP.Producto "
     End If
     sSQL = sSQL & "WHERE DF.Periodo = CP.Periodo " _
          & "AND DF.Item = CP.Item " _
          & "AND DF.Codigo = CP.Codigo_Inv " _
          & "AND DF.Producto = '" & Ninguno & "' "
     Ejecutar_SQL_SP sSQL
     DGFactura.Visible = True
     RatonNormal
  End If
  Unload Me
End Sub

Private Sub Command2_Click()
   Control_Procesos Normal, "Salir de Puntos de Venta"
   Unload Me
End Sub

Private Sub Command3_Click()
  Caja_Pendiente_PV
End Sub

Private Sub Command4_Click()
  If OpcIndividual.value Then
     FA.Factura = Val(DCTicket)
     If Grafico_PV Then
        Imprimir_Punto_Venta_Grafico FA
     Else
        Imprimir_Punto_Venta FA
     End If
  Else
     MsgBox "En esta opcion no se puede imprimir reporte"
  End If
  Caja_Pendiente_PV
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_LostFocus()
   CodigoCli = "9999999999"
   With AdoCliente.Recordset
    If .RecordCount > 0 Then
       .MoveFirst
       .Find ("Cliente = '" & DCCliente.Text & "' ")
        If Not .EOF Then CodigoCli = .Fields("Codigo")
    End If
   End With
End Sub

Private Sub DCLinea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCLinea_LostFocus()
  FA.Cod_CxC = DCLinea
  Lineas_De_CxC FA
  FA.Nuevo_Doc = True
  FA.Factura = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
  Label4.Caption = " No. " & FA.Serie & "-" & Format$(FA.Factura, "00000000")
End Sub

Private Sub Form_Activate()
  DGFactura.Visible = False
  MBFechaI = FechaSistema
  MBFechaF = FechaSistema
  sSQL = "SELECT Codigo & Space(5) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE MidStrg(Codigo,1,1) = '1' " _
       & "AND TC IN ('CJ','BA') " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Lineas " _
       & "WHERE TL <> " & Val(adFalse) & " " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fact <> 'PV' " _
       & "ORDER BY Codigo "
  SelectDB_Combo DCLinea, AdoLinea, sSQL, "Concepto"
  
  sSQL = "SELECT Ticket " _
       & "FROM Trans_Ticket " _
       & "WHERE TC = 'PV' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND C = " & Val(adFalse) & " " _
       & "GROUP BY Ticket " _
       & "ORDER BY Ticket DESC "
  SelectDB_Combo DCTicket, AdoTicket, sSQL, "Ticket"
  
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' "
  SelectDB_Combo DCCliente, AdoCliente, sSQL, "Cliente"
  DCCliente.Text = "CONSUMIDOR FINAL"
  RatonNormal
  CierreDelDia
  Caja_Pendiente_PV
End Sub

Private Sub Form_Load()
   'CentrarForm FConvertirPV
   ConectarAdodc AdoBanco
   ConectarAdodc AdoLinea
   ConectarAdodc AdoTicket
   ConectarAdodc AdoFactura
   ConectarAdodc AdoIngCaja
   ConectarAdodc AdoCliente
   Encerar_Factura FA
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  MBFechaF = MBFechaI
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Public Sub Caja_Pendiente_PV()
  DGFactura.Visible = False
  RatonReloj
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  Numero = Val(DCTicket.Text)
  If OpcResumen.value Then
     sSQL = "SELECT Fecha,CodigoC,Codigo_Inv,SUM(Cantidad) As Cant,SUM(Descuento) As TDescuento,(SUM(Total)/SUM(Cantidad)) As PVP,SUM(Total) As Total_PV "
  Else
     sSQL = "SELECT Fecha,Ticket,CodigoC,Codigo_Inv,Cantidad As Cant,Precio As PVP,Total As Total_PV "
  End If
  sSQL = sSQL & "FROM Trans_Ticket " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'PV' " _
       & "AND C = " & Val(adFalse) & " "
  If OpcIndividual.value Then sSQL = sSQL & "AND Ticket = " & Numero & " "
  If OpcResumen.value Then
     sSQL = sSQL & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "GROUP BY Fecha,CodigoC,Codigo_Inv "
  End If
  sSQL = sSQL & "ORDER BY Fecha,CodigoC,Codigo_Inv "
  Select_Adodc_Grid DGFactura, AdoFactura, sSQL
  RatonReloj
  Total = 0: Contador = 0
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          FConvertirPV.Caption = Format$(Contador / .RecordCount, "00%")
          Total = Total + .Fields("Total_PV")
         .MoveNext
       Loop
   End If
  End With
  LabelSaldo.Caption = Format$(Total, "#,##0.00")
  DGFactura.Visible = True
  RatonNormal
End Sub

Private Sub OpcIndividual_Click()
  Label3.Visible = True
  DCCliente.Visible = True
  Label1.Visible = True
  DCBanco.Visible = True
End Sub

Private Sub OpcResumen_Click()
  Label3.Visible = False
  DCCliente.Visible = False
  Label1.Visible = False
  DCBanco.Visible = False
End Sub
