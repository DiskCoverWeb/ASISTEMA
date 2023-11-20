VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacturaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OPERACION DE CREDITO"
   ClientHeight    =   7065
   ClientLeft      =   5025
   ClientTop       =   4380
   ClientWidth     =   7470
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
   ScaleHeight     =   7065
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtMonto 
      Height          =   330
      Left            =   105
      MaxLength       =   8
      TabIndex        =   10
      Text            =   "0"
      Top             =   1680
      Width           =   1905
   End
   Begin VB.TextBox TxtInt 
      Height          =   330
      Left            =   3780
      MaxLength       =   8
      TabIndex        =   21
      Text            =   "0"
      Top             =   3150
      Width           =   750
   End
   Begin MSDataGridLib.DataGrid DGAbonos 
      Bindings        =   "FCredito.frx":0000
      Height          =   3060
      Left            =   105
      TabIndex        =   26
      Top             =   3570
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   5398
      _Version        =   393216
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin VB.TextBox TextCheqNo 
      Height          =   330
      Left            =   3045
      MaxLength       =   8
      TabIndex        =   19
      Text            =   "1"
      Top             =   3150
      Width           =   750
   End
   Begin MSAdodcLib.Adodc AdoIngCaja 
      Height          =   330
      Left            =   6405
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
      Left            =   6405
      Top             =   1995
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
   Begin VB.TextBox TextCajaMN 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   4515
      MaxLength       =   14
      TabIndex        =   23
      Text            =   "0"
      Top             =   3150
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   6405
      Picture         =   "FCredito.frx":0018
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1050
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   6405
      Picture         =   "FCredito.frx":08E2
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   105
      Width           =   960
   End
   Begin MSDataListLib.DataCombo DCFactura 
      Bindings        =   "FCredito.frx":0D24
      DataSource      =   "AdoFactura"
      Height          =   315
      Left            =   4200
      TabIndex        =   5
      Top             =   420
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Factura"
   End
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   15
      Top             =   3150
      Width           =   1380
      _ExtentX        =   2434
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
      Left            =   6405
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
   Begin MSDataListLib.DataCombo DCTipo 
      Bindings        =   "FCredito.frx":0D3D
      DataSource      =   "AdoDetAcomp"
      Height          =   315
      Left            =   3255
      TabIndex        =   3
      Top             =   420
      Width           =   855
      _ExtentX        =   1508
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
   Begin MSAdodcLib.Adodc AdoDetAcomp 
      Height          =   330
      Left            =   6405
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
   Begin MSAdodcLib.Adodc AdoAbonos 
      Height          =   330
      Left            =   105
      Top             =   6615
      Width           =   6210
      _ExtentX        =   10954
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
      Caption         =   "Abonos"
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
   Begin MSDataListLib.DataCombo DCSerie 
      Bindings        =   "FCredito.frx":0D57
      DataSource      =   "AdoSerie"
      Height          =   360
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "001001"
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
   Begin MSAdodcLib.Adodc AdoSerie 
      Height          =   330
      Left            =   6405
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
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SERIE"
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3060
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Interés"
      Height          =   330
      Left            =   3780
      TabIndex        =   20
      Top             =   2835
      Width           =   750
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3675
      TabIndex        =   11
      Top             =   1680
      Width           =   2640
   End
   Begin VB.Label LblNota 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nota"
      ForeColor       =   &H00C000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   2415
      Width           =   6210
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TIPO"
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3255
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   2205
      TabIndex        =   7
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA DE EMISION"
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   840
      Width           =   2115
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Meses"
      Height          =   330
      Left            =   3045
      TabIndex        =   18
      Top             =   2835
      Width           =   750
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Fecha Abonos"
      Height          =   330
      Left            =   105
      TabIndex        =   14
      Top             =   2835
      Width           =   1380
   End
   Begin VB.Label LblObs 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
      ForeColor       =   &H00C000C0&
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   2100
      Width           =   6210
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3570
      TabIndex        =   8
      Top             =   840
      Width           =   2745
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dividendo Mensual"
      Height          =   330
      Left            =   4515
      TabIndex        =   22
      Top             =   2835
      Width           =   1800
   End
   Begin VB.Label LabelSaldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   1470
      TabIndex        =   17
      Top             =   3150
      Width           =   1590
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F&actura No."
      Height          =   330
      Left            =   4200
      TabIndex        =   4
      Top             =   105
      Width           =   2115
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Saldo Pendiente"
      Height          =   330
      Left            =   1470
      TabIndex        =   16
      Top             =   2835
      Width           =   1590
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   9
      Top             =   1260
      Width           =   6210
   End
End
Attribute VB_Name = "FacturaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Listar_Facturas_Pendientes()
  TipoFactura = DCTipo
  Autorizacion = SinEspaciosDer(DCSerie)
  SerieFactura = SinEspaciosIzq(DCSerie)
  
  SQL1 = "SELECT F.TC,F.Factura,F.CodigoC,F.Fecha,F.Fecha_V,F.Saldo_MN,F.Cta_CxP,F.Nota," _
       & "F.Observacion,C.Cliente,C.Direccion,C.CI_RUC,C.Telefono,C.Grupo,F.Autorizacion,F.Serie " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.T = '" & Pendiente & "' " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.TC = '" & TipoFactura & "' " _
       & "AND F.Autorizacion = '" & Autorizacion & "' " _
       & "AND F.Serie = '" & SerieFactura & "' " _
       & "AND F.Saldo_MN > 0 " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.TC,F.Factura "
  If TipoFactura = "NV" Then Label2.Caption = "Nota de Venta No." Else Label2.Caption = "Factura No."
  SelectDBCombo DCFactura, AdoFactura, SQL1, "Factura"
End Sub

Private Sub Command1_Click()
  FechaValida MBFecha
  Mensajes = "Esta Seguro que desea grabar Crédito"
  Titulo = "Formulario de Grabación."
  If BoxMensaje = vbYes Then
     sSQL = "DELETE * " _
          & "FROM Trans_Prestamos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Credito_No = '" & Format$(TA.Factura, "0000000") & "' " _
          & "AND Autorizacion = '" & Autorizacion & "' " _
          & "AND Serie = '" & SerieFactura & "' " _
          & "AND Cuenta_No = '" & TA.CodigoC & "' " _
          & "AND TP = '" & TA.TP & "' "
     ConectarAdoExecute sSQL

     sSQL = "SELECT * " _
          & "FROM Asiento_P " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' " _
          & "ORDER BY Cuotas "
     SelectAdodc AdoAbonos, sSQL
     With AdoAbonos.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             If .Fields("Pagos") <> 0 Then
                 SetAdoAddNew "Trans_Prestamos"
                 SetAdoFields "T", Normal
                 SetAdoFields "TP", TA.TP
                 SetAdoFields "Serie", TA.Serie
                 SetAdoFields "Autorizacion", TA.Autorizacion
                 SetAdoFields "Fecha", .Fields("Fecha")
                 SetAdoFields "Cuota_No", .Fields("T_No")
                 SetAdoFields "Credito_No", Format$(TA.Factura, "0000000")
                 SetAdoFields "Cuenta_No", TA.CodigoC
                 SetAdoFields "Pagos", .Fields("Pagos")
                 SetAdoFields "Saldo", .Fields("Saldo")
                 SetAdoFields "Interes", .Fields("Interes")
                 SetAdoFields "Cta", TA.Cta_CxP
                 SetAdoFields "Item", NumEmpresa
                 SetAdoFields "CodigoU", CodigoUsuario
                 SetAdoUpdate
             End If
            .MoveNext
          Loop
      End If
     End With
     sSQL = "DELETE * " _
          & "FROM Asiento_P " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     ConectarAdoExecute sSQL
     TxtMonto = "0.00"
     DCFactura.SetFocus
  End If
  'Unload AbonoEfectivo
End Sub

Private Sub Command2_Click()
   Control_Procesos Normal, "Salir de Factura a Crédito"
   Unload Me
End Sub

Private Sub DCBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCBanco_LostFocus()
  TextCajaMN.SetFocus
End Sub

Private Sub DCFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCFactura_LostFocus()
  Codigo1 = Ninguno
  Saldo = 0
  Total_IVA = 0
  Total_Ret = 0
  TotalCajaMN = 0
  TotalCajaME = 0
  Total_Bancos = 0
  Total_Tarjeta = 0
  Cotizacion = 0
  TotalDolar = 0
  Saldo_ME = 0
  Label3.Caption = ""
  Label1.Caption = ""
  LabelSaldo.Caption = ""
  With AdoFactura.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Factura = " & Val(DCFactura) & " ")
       If Not .EOF Then
          Label8.Caption = " " & .Fields("Fecha")
          Label11.Caption = "C.I./R.U.C.: " & .Fields("CI_RUC")
          Grupo_No = " " & .Fields("Grupo")
          TextCheqNo = "1"
          LblObs.Caption = " " & .Fields("Observacion")
          LblNota.Caption = " " & .Fields("Nota")
          CodigoCliente = .Fields("CodigoC")
          NombreCliente = .Fields("Cliente")
          DireccionCli = .Fields("Direccion")
          Factura_No = .Fields("Factura")
          Cta_Cobrar = .Fields("Cta_CxP")
          TipoFactura = .Fields("TC")
          Saldo = .Fields("Saldo_MN")
         'Datos del Abonos
          TA.Serie = .Fields("Serie")
          TA.Autorizacion = .Fields("Autorizacion")
          TA.TP = TipoFactura
          TA.Fecha = MBFecha
          TA.Cta_CxP = Cta_Cobrar
          TA.Factura = Factura_No
          TA.CodigoC = CodigoCliente
          TxtMonto = Format$(Saldo, "#,##0.00")
          LabelSaldo.Caption = Format$(Saldo, "#,##0.00")
          Command1.Enabled = True
          Label3.Caption = NombreCliente
          Label1.Caption = "Factura No. " & Format$(Factura_No, "0000000")
          SaldoDisp = Saldo - TotalCajaMN - TotalCajaME - Total_Bancos - Total_Tarjeta - Total_IVA - Total_Ret
          MBFecha.SetFocus
       Else
          MsgBox "Esta Factura no esta pendiente"
          Command1.Enabled = False
          DCTipo.SetFocus
       End If
    End If
  End With
End Sub

Private Sub DCSerie_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCSerie_LostFocus()
  Autorizacion = SinEspaciosDer(DCSerie)
  SerieFactura = SinEspaciosIzq(DCSerie)
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "GROUP BY TC " _
       & "ORDER BY TC DESC "
  SelectDBCombo DCTipo, AdoDetAcomp, sSQL, "TC"
End Sub

Private Sub DCTipo_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCTipo_LostFocus()
   TipoFactura = DCTipo.Text
   If TipoFactura = "" Then TipoFactura = "FA"
   Listar_Facturas_Pendientes
End Sub

Private Sub Form_Activate()
  ControlEsNumerico TextCajaMN
  
  sSQL = "SELECT Vencimiento,(Serie & ' - ' & Autorizacion) As SerieFact " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "GROUP BY Vencimiento,Serie,Autorizacion " _
       & "ORDER BY Vencimiento DESC,Autorizacion,Serie "
  SelectDBCombo DCSerie, AdoSerie, sSQL, "SerieFact"

  Autorizacion = SinEspaciosDer(DCSerie)
  SerieFactura = SinEspaciosIzq(DCSerie)
  sSQL = "SELECT TC " _
       & "FROM Facturas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Autorizacion = '" & Autorizacion & "' " _
       & "AND Serie = '" & SerieFactura & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "GROUP BY TC " _
       & "ORDER BY TC DESC "
  SelectDBCombo DCTipo, AdoDetAcomp, sSQL, "TC"
  
  sSQL = "SELECT Cuotas,TP,Fecha,Pagos,Saldo,T_No " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY Fecha,Cuotas "
  SelectDataGrid DGAbonos, AdoAbonos, sSQL
  
  DiarioCaja = ReadSetDataNum("Recibo_No", True, False)
  Trans_No = 100
  Mifecha = BuscarFecha(FechaTexto)
  MBFecha.Text = FechaSistema
  RatonNormal
  DCSerie.SetFocus
End Sub

Private Sub Form_Load()
   CentrarForm FacturaCredito
   ConectarAdodc AdoSerie
   ConectarAdodc AdoBanco
   ConectarAdodc AdoAbonos
   ConectarAdodc AdoFactura
   ConectarAdodc AdoIngCaja
   ConectarAdodc AdoDetAcomp
  
   sSQL = "DELETE * " _
        & "FROM Asiento_P " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
End Sub

Private Sub TextCajaMN_GotFocus()
  MarcarTexto TextCajaMN
End Sub

Private Sub TextCajaMN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextCajaMN_LostFocus()
  TextoValido TextCajaMN, True
  Mifecha = MBFecha
  SaldoDisp = Saldo
  
  'MsgBox ".........."
  Generar_Tabla_Prestamo_Sobre_Saldos MBFecha, AdoAbonos, DGAbonos, TxtInt, TextCheqNo, TxtMonto, False, TipoFactura, False
  
  
''''  sSQL = "DELETE * " _
''''       & "FROM Asiento_P " _
''''       & "WHERE Item = '" & NumEmpresa & "' " _
''''       & "AND CodigoU = '" & CodigoUsuario & "' "
''''  ConectarAdoExecute sSQL
''''      SetAdoAddNew "Asiento_P"
''''      SetAdoFields "TP", TipoFactura
''''      SetAdoFields "Fecha", Mifecha
''''      SetAdoFields "Pagos", 0
''''      SetAdoFields "Cuotas", 0
''''      SetAdoFields "Saldo", Saldo
''''      SetAdoFields "T_No", Trans_No
''''      SetAdoFields "Item", NumEmpresa
''''      SetAdoFields "CodigoU", CodigoUsuario
''''      SetAdoUpdate
''''  SaldoDisp = SaldoDisp - Val(TextCajaMN)
''''  Mifecha = SiguienteMes(Mifecha)
''''  For I = 1 To Val(TextCheqNo)
''''      If SaldoDisp < 0 Then
''''         Abono = Val(TextCajaMN) + SaldoDisp
''''         SaldoDisp = 0
''''      Else
''''         Abono = Val(TextCajaMN)
''''      End If
''''      SetAdoAddNew "Asiento_P"
''''      SetAdoFields "TP", TipoFactura
''''      SetAdoFields "Fecha", Mifecha
''''      SetAdoFields "Pagos", Abono
''''      SetAdoFields "Cuotas", I
''''      SetAdoFields "Saldo", SaldoDisp
''''      SetAdoFields "T_No", Trans_No
''''      SetAdoFields "Item", NumEmpresa
''''      SetAdoFields "CodigoU", CodigoUsuario
''''      SetAdoUpdate
''''      Mifecha = SiguienteMes(Mifecha)
''''      SaldoDisp = SaldoDisp - Val(TextCajaMN)
''''  Next
  sSQL = "SELECT Fecha,Interes,Capital,Pagos,Saldo,T_No,TP " _
       & "FROM Asiento_P " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "ORDER BY Fecha,Cuotas "
  SelectDataGrid DGAbonos, AdoAbonos, sSQL
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
End Sub

Private Sub TextCheqNo_GotFocus()
  MarcarTexto TextCheqNo
End Sub

Private Sub TextCheqNo_LostFocus()
  TextoValido TextCheqNo
  If Val(TextCheqNo) <= 0 Then TextCheqNo = "1"
  TextCajaMN = Format$(Saldo / Val(TextCheqNo), "#,##0.00")
End Sub

Private Sub TxtRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRecibo_LostFocus()
  TxtRecibo = Format$(Val(TxtRecibo), "0000000")
End Sub

Private Sub TxtInt_GotFocus()
  MarcarTexto TxtInt
End Sub

Private Sub TxtInt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtInt_LostFocus()
  TextoValido TxtInt, True
End Sub
