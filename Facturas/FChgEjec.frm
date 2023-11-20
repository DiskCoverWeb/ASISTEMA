VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FCambioEjecutivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adulto"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12750
   Icon            =   "FChgEjec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Cuota"
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
      Left            =   9555
      Picture         =   "FChgEjec.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   945
      Width           =   960
   End
   Begin VB.TextBox TxtCuota 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7350
      MaxLength       =   14
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "FChgEjec.frx":0A14
      Top             =   420
      Width           =   2115
   End
   Begin VB.OptionButton OpcFA 
      Caption         =   "Seleccionar Por Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9555
      TabIndex        =   5
      Top             =   525
      Width           =   2430
   End
   Begin VB.OptionButton OpcCli 
      Caption         =   "Seleccionar Por Cliente"
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
      Left            =   9555
      TabIndex        =   4
      Top             =   105
      Value           =   -1  'True
      Width           =   2325
   End
   Begin MSDataListLib.DataCombo DCEjecNuevo 
      Bindings        =   "FChgEjec.frx":0A19
      DataSource      =   "AdoEjecutivos"
      Height          =   345
      Left            =   2940
      TabIndex        =   9
      Top             =   1470
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCZonaActual 
      Bindings        =   "FChgEjec.frx":0A35
      DataSource      =   "AdoZonaAct"
      Height          =   345
      Left            =   2100
      TabIndex        =   7
      Top             =   945
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox LstClientes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5730
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   1995
      Width           =   12510
   End
   Begin MSDataListLib.DataCombo DCEjecActual 
      Bindings        =   "FChgEjec.frx":0A4E
      DataSource      =   "AdoEjecutivos"
      Height          =   345
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   609
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   10605
      Picture         =   "FChgEjec.frx":0A6A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   945
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
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
      Left            =   11655
      Picture         =   "FChgEjec.frx":1334
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   945
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoEjecutivos 
      Height          =   330
      Left            =   105
      Top             =   7665
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
      Caption         =   "Ejecutivos"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   2415
      Top             =   7665
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoZonaAct 
      Height          =   330
      Left            =   6405
      Top             =   7665
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "ZonasAct"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   4515
      Top             =   7665
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CUOTA MENSUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   7350
      TabIndex        =   2
      Top             =   105
      Width           =   2115
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CAMBIAR POR EL EJECUTVO"
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
      TabIndex        =   8
      Top             =   1470
      Width           =   2850
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ZONAS ACTUALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   945
      Width           =   2010
   End
   Begin VB.Label LblCliente 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EJECUTIVO ACTUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7155
   End
End
Attribute VB_Name = "FCambioEjecutivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cta_Zona As String

Private Sub Command1_Click()
    Mensajes = "Desea cambiar ejecutivo de Venta:" & vbCrLf _
             & DCEjecActual & vbCrLf _
             & "del Grupo: " & DCZonaActual & vbCrLf _
             & "Por el nuevo Ejecutivo:" & vbCrLf _
             & DCEjecNuevo
    Titulo = "CAMBIOS DE EJECUTIVOS DE VENTA"
    If BoxMensaje = vbYes Then
       RatonReloj
       For I = 2 To LstClientes.ListCount - 1
           If LstClientes.Selected(I) Then
              CodigoCli = SinEspaciosDer(LstClientes.List(I))
              sSQL = "SELECT Codigo, Cod_Ejec, Cta_CxP, FA, ID " _
                   & "FROM Clientes " _
                   & "WHERE Codigo = '" & CodigoCli & "' "
              Select_Adodc AdoClientes, sSQL
              With AdoClientes.Recordset
               If .RecordCount > 0 Then
                  .Fields("Cod_Ejec") = CodigoVen
                  .Fields("Cta_CxP") = Cta_Zona
                  .Fields("FA") = True
                  .Update
               End If
              End With
              sSQL = "UPDATE Facturas " _
                   & "SET Cod_Ejec = '" & CodigoVen & "' " _
                   & "WHERE Item = '" & NumEmpresa & "' " _
                   & "AND Periodo = '" & Periodo_Contable & "' " _
                   & "AND CodigoC = '" & CodigoCli & "' "
              Ejecutar_SQL_SP sSQL
           End If
       Next I
    
'''          sSQL = "UPDATE Clientes " _
'''               & "SET Cod_Ejec = '" & CodigoVen & "' " _
'''               & "WHERE Cod_Ejec = '" & CodigoEjecutivo & "' " _
'''               & "AND Cta_CxP = '" & Cta_Zona & "' "
'''          Ejecutar_SQL_SP sSQL
       
       If SQL_Server Then
          sSQL = "UPDATE Detalle_Factura " _
               & "SET Cod_Ejec = F.Cod_Ejec " _
               & "FROM Detalle_Factura As DF, Facturas As F "
       Else
          sSQL = "UPDATE Detalle_Factura As DF, Facturas As F " _
               & "SET DF.Cod_Ejec = F.Cod_Ejec "
       End If
       sSQL = sSQL _
            & "WHERE F.Item = '" & NumEmpresa & "' " _
            & "AND F.Periodo = '" & Periodo_Contable & "' " _
            & "AND F.Cta_CxP = '" & Cta_Zona & "' " _
            & "AND DF.Item = F.Item " _
            & "AND DF.Periodo = F.Periodo " _
            & "AND DF.TC = F.TC " _
            & "AND DF.Serie = F.Serie " _
            & "AND DF.Factura = F.Factura " _
            & "AND DF.Autorizacion = F.Autorizacion "
       Ejecutar_SQL_SP sSQL
       
       If SQL_Server Then
          sSQL = "UPDATE Trans_Abonos " _
               & "SET Cod_Ejec = F.Cod_Ejec " _
               & "FROM Trans_Abonos As DF, Facturas As F "
       Else
          sSQL = "UPDATE Trans_Abonos As DF, Facturas As F " _
               & "SET DF.Cod_Ejec = F.Cod_Ejec "
       End If
       sSQL = sSQL _
            & "WHERE F.Item = '" & NumEmpresa & "' " _
            & "AND F.Periodo = '" & Periodo_Contable & "' " _
            & "AND F.Cta_CxP = '" & Cta_Zona & "' " _
            & "AND DF.Item = F.Item " _
            & "AND DF.Periodo = F.Periodo " _
            & "AND DF.TP = F.TC " _
            & "AND DF.Serie = F.Serie " _
            & "AND DF.Factura = F.Factura " _
            & "AND DF.Autorizacion = F.Autorizacion "
       Ejecutar_SQL_SP sSQL
      
       sSQL = "SELECT C.Cta_CxP, CC.Cuenta, COUNT(C.Codigo) " _
            & "FROM Catalogo_Cuentas As CC, Clientes As C " _
            & "WHERE CC.Item = '" & NumEmpresa & "' " _
            & "AND CC.Periodo = '" & Periodo_Contable & "' " _
            & "AND CC.Codigo = C.Cta_CxP " _
            & "GROUP BY C.Cta_CxP, CC.Cuenta " _
            & "ORDER BY CC.Cuenta "
       Select_Adodc AdoZonaAct, sSQL
       With AdoZonaAct.Recordset
        If .RecordCount > 0 Then
            Do While Not .EOF
               Codigo1 = .Fields("Cta_CxP")
               Codigo2 = .Fields("Cuenta")
               I = InStr(Codigo2, "ZONA")
               If I > 0 Then
                  Codigo3 = TrimStrg(MidStrg(Codigo2, I, 10))
                  sSQL = "UPDATE Clientes " _
                       & "SET Grupo = '" & Codigo3 & "' " _
                       & "WHERE Cta_CxP = '" & Codigo1 & "' "
                  Ejecutar_SQL_SP sSQL
               End If
              .MoveNext
            Loop
        End If
       End With
       RatonNormal
      'MsgBox CodigoEjecutivo & vbCrLf & CodigoVen & vbCrLf & sSQL
       Unload FCambioEjecutivo
    End If
End Sub

Private Sub Command2_Click()
  Unload FCambioEjecutivo
End Sub

Private Sub Command3_Click()
    Mensajes = "Desea cambiar ejecutivo de Venta:" & vbCrLf _
             & DCEjecActual & vbCrLf _
             & "del Grupo: " & DCZonaActual & vbCrLf _
             & "Por el nuevo Ejecutivo:" & vbCrLf _
             & DCEjecNuevo
    Titulo = "CAMBIOS DE EJECUTIVOS DE VENTA"
    If BoxMensaje = vbYes Then
       RatonReloj
       sSQL = "UPDATE Accesos " _
            & "SET Cuota_Venta = " & Val(TxtCuota) & "" _
            & "WHERE Codigo = '" & CodigoEjecutivo & "' "
       Ejecutar_SQL_SP sSQL
       RatonNormal
    Else
       RatonNormal
    End If
End Sub

Private Sub DCEjecActual_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCEjecActual_LostFocus()
    CodigoEjecutivo = Leer_Ejecutivo(DCEjecActual)
    sSQL = "SELECT Cuota_Venta " _
         & "FROM  Accesos " _
         & "WHERE Codigo = '" & CodigoEjecutivo & "' "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then TxtCuota = Format(.Fields("Cuota_Venta"), "#,##0.00") Else TxtCuota = "0.00"
    End With
    Listar_Zonas_Ejecutivo CodigoEjecutivo
End Sub

Private Sub DCEjecNuevo_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub DCEjecNuevo_LostFocus()
   CodigoVen = Leer_Ejecutivo(DCEjecNuevo)
End Sub

Private Sub DCZonaActual_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCZonaActual_LostFocus()
   Grupo_No = Leer_Zona(DCZonaActual)
   Clientes_x_Ejecutivo CodigoEjecutivo, Grupo_No
End Sub

Private Sub Form_Activate()
   'MsgBox CodigoEjecutivo & vbCrLf & Cta
'''    If SQL_Server Then
'''       sSQL = "UPDATE Clientes " _
'''            & "SET Cod_Ejec = F.Cod_Ejec, Cta_CxP = F.Cta_CxP " _
'''            & "FROM Clientes As C,Facturas As F "
'''    Else
'''       sSQL = "UPDATE Clientes As C,Facturas As F " _
'''            & "SET C.Cod_Ejec = F.Cod_Ejec, C.Cta_CxP = F.Cta_CxP "
'''    End If
'''    sSQL = sSQL _
'''         & "WHERE F.Item = '" & NumEmpresa & "' " _
'''         & "AND F.Periodo = '" & Periodo_Contable & "' " _
'''         & "AND F.Fecha >= '20180101' " _
'''         & "AND C.Codigo = F.CodigoC "
'''    Ejecutar_SQL_SP sSQL
   
    sSQL = "SELECT CR.Codigo,C.Cliente,C.Cta_CxP,CR.Porc_Com " _
         & "FROM Catalogo_Rol_Pagos As CR, Clientes As C " _
         & "WHERE CR.Item = '" & NumEmpresa & "' " _
         & "AND CR.Periodo = '" & Periodo_Contable & "' " _
         & "AND CR.Codigo = C.Codigo " _
         & "ORDER BY C.Cliente "
    SelectDB_Combo DCEjecActual, AdoEjecutivos, sSQL, "Cliente"
    SelectDB_Combo DCEjecNuevo, AdoEjecutivos, sSQL, "Cliente"
    If CodigoEjecutivo = "" Then CodigoEjecutivo = Ninguno
    DCEjecActual = CodigoEjecutivo
    CodigoEjecutivo = Leer_Ejecutivo(DCEjecActual)
    Listar_Zonas_Ejecutivo CodigoEjecutivo
    RatonNormal
    DCEjecActual.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm FCambioEjecutivo
  ConectarAdodc AdoAux
  ConectarAdodc AdoEjecutivos
  ConectarAdodc AdoZonaAct
  ConectarAdodc AdoClientes
  FCambioEjecutivo.Caption = "CAMBIO DE EJECUTIVOS DE VENTAS EN GRUPO"
End Sub

Public Sub Listar_Zonas_Ejecutivo(Cod_Ejec As String)
   '& "AND C.Cod_Ejec = '" & Cod_Ejec & "' "
    sSQL = "SELECT C.Cta_CxP, CC.Cuenta, COUNT(C.Codigo) " _
         & "FROM Catalogo_Cuentas As CC, Clientes As C " _
         & "WHERE CC.Item = '" & NumEmpresa & "' " _
         & "AND CC.Periodo = '" & Periodo_Contable & "' " _
         & "AND CC.Codigo = C.Cta_CxP " _
         & "GROUP BY C.Cta_CxP, CC.Cuenta " _
         & "ORDER BY CC.Cuenta "
    SelectDB_Combo DCZonaActual, AdoZonaAct, sSQL, "Cuenta"
    'MsgBox sSQL
End Sub

Public Function Leer_Ejecutivo(Ejecutivo) As String
Dim CodEjec As String
    CodEjec = Ninguno
    If Ejecutivo = "" Then Ejecutivo = Ninguno
    With AdoEjecutivos.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Cliente = '" & Ejecutivo & "' ")
         If Not .EOF Then
            CodEjec = .Fields("Codigo")
            Cta_Aux = .Fields("Cta_CxP")
         End If
     End If
    End With
    Leer_Ejecutivo = CodEjec
End Function

Public Function Leer_Zona(Zona As String) As String
Dim SubCad As String
Dim GZona As String
    Cta_Zona = Ninguno
    GZona = Ninguno
    If Zona = "" Then Zona = Ninguno
    With AdoZonaAct.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Cuenta = '" & Zona & "' ")
         If Not .EOF Then
            Cta_Zona = .Fields("Cta_CxP")
            SubCad = .Fields("Cuenta")
            I = InStr(SubCad, "ZONA")
            If I > 0 Then GZona = TrimStrg(MidStrg(SubCad, I, 10))
         End If
     End If
    End With
    Leer_Zona = GZona
End Function

Public Sub Clientes_x_Ejecutivo(CodEjec As String, CtaZona As String)
    LstClientes.Visible = False
    LstClientes.Clear
   'MsgBox CtaZona
    If OpcCli.value Then
       sSQL = "SELECT Codigo, Cliente " _
            & "FROM Clientes " _
            & "WHERE Cod_Ejec = '" & CodEjec & "' " _
            & "AND Grupo = '" & CtaZona & "' " _
            & "ORDER BY Cliente "
       Select_Adodc AdoClientes, sSQL
    Else
       sSQL = "SELECT C.Codigo, C.Cliente,COUNT(F.Factura) " _
            & "FROM Clientes As C, Facturas As F " _
            & "WHERE F.Cod_Ejec = '" & CodEjec & "' " _
            & "AND F.Item = '" & NumEmpresa & "' " _
            & "AND F.Periodo = '" & Periodo_Contable & "' " _
            & "AND C.Grupo = '" & CtaZona & "' " _
            & "AND C.Codigo = F.CodigoC " _
            & "GROUP BY C.Codigo, C.Cliente " _
            & "ORDER BY Cliente "
       Select_Adodc AdoClientes, sSQL
    End If
    With AdoClientes.Recordset
     If .RecordCount > 0 Then
         Cadena = DCZonaActual & " - " & DCEjecActual
         LstClientes.AddItem Cadena & String(80 - Len(Cadena), " ") & "CODIGO (" & .RecordCount & ")"
         LstClientes.AddItem String(95, "-")
         Do While Not .EOF
            LstClientes.AddItem .Fields("Cliente") & String(80 - Len(.Fields("Cliente")), " ") & .Fields("Codigo")
           .MoveNext
         Loop
     End If
    End With
    For I = 2 To LstClientes.ListCount - 1
        LstClientes.Selected(I) = True
    Next I
    LstClientes.ListIndex = 0
    LstClientes.Visible = True
End Sub

Private Sub LstClientes_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TextAux As String
    Keys_Especiales Shift
    If CtrlDown And KeyCode = vbKeyP Then
       TextAux = ""
       For I = 0 To LstClientes.ListCount - 1
           TextAux = TextAux & vbTab & LstClientes.List(I) & vbCrLf
       Next I
       Imprimir_Texo_Impresora TextAux
    End If
End Sub

Private Sub TxtCuota_GotFocus()
  MarcarTexto TxtCuota
End Sub

Private Sub TxtCuota_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCuota_LostFocus()
  TextoValido TxtCuota, True, , 2
End Sub
