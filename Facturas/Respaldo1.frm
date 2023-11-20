VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Respaldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Espere un momento....     Estoy procesando las bases"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "Respaldo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   6840
      Begin VB.TextBox TextUnidad 
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
         Left            =   1365
         MaxLength       =   1
         TabIndex        =   6
         Text            =   "A"
         Top             =   630
         Width           =   1275
      End
      Begin MSMask.MaskEdBox MBoxFechaF 
         Height          =   330
         Left            =   4725
         TabIndex        =   4
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   327680
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
      Begin MSMask.MaskEdBox MBoxFechaI 
         Height          =   330
         Left            =   2625
         TabIndex        =   2
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   327680
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &Unidad:"
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
         TabIndex        =   5
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta"
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
         Left            =   3990
         TabIndex        =   3
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha de Respaldo Desde:"
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
         TabIndex        =   1
         Top             =   210
         Width           =   2535
      End
   End
   Begin VB.Data DataEmpresas 
      Caption         =   "Empresas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3780
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   210
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir de Respaldos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   4725
      Picture         =   "Respaldo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2835
      Width           =   2220
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Act&ualizar Cancelados de las Sucursales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   2415
      Picture         =   "Respaldo.frx":06C4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2835
      Width           =   2220
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Actualizar Transacciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   4725
      Picture         =   "Respaldo.frx":0B06
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   2220
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Primero &Descomprima para actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   2415
      Picture         =   "Respaldo.frx":0F48
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1680
      Width           =   2220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Transacciones del Día"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      Picture         =   "Respaldo.frx":138A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   2220
   End
   Begin VB.Data DataOld 
      Caption         =   "Old"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   525
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataAux 
      Caption         =   "Aux"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3780
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   525
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataAct 
      Caption         =   "Act"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   525
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   1995
      TabIndex        =   13
      Top             =   210
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   210
      TabIndex        =   12
      Top             =   210
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DataQuery 
      Caption         =   "Query"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5565
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   210
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   105
      TabIndex        =   7
      Top             =   1155
      Width           =   6840
   End
End
Attribute VB_Name = "Respaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 'Envios pendientes matriz
  RatonReloj
  'MsgBox RutaBackup
  Respaldos.Caption = "Copiando bases requeridas"
  Dir1.Path = RutaEmpresa
  File1.filename = Dir1.Path & "\PRODUCC*.MDB"
  For Ind = 0 To File1.ListCount - 1
      RutaOrigen = Dir1.Path & "\" & File1.List(Ind)
      RutaDestino = RutaBackup & "\" & File1.List(Ind)
      Respaldos.Caption = "Copiando la base: " & RutaDestino
      FileCopy RutaOrigen, RutaDestino
  Next Ind
  RutaOrigen = RutaSistema & "\Empresas.mdb"
  RutaDestino = RutaBackup & "\Empresas.mdb"
  FileCopy RutaOrigen, RutaDestino
  RutaOrigen = RutaSistema & "\pkzip.exe"
  RutaDestino = RutaBackup & "\pkzip.exe"
  FileCopy RutaOrigen, RutaDestino
  RutaOrigen = RutaSistema & "\pkunzip.exe"
  RutaDestino = RutaBackup & "\pkunzip.exe"
  FileCopy RutaOrigen, RutaDestino
  Dir1.Path = RutaBackup
  File1.filename = Dir1.Path & "\*.MDB"
  DataQuery.DatabaseName = RutaBackup & "\PRODUCC.MDB"
  DataQuery.Refresh
  DataAct.DatabaseName = RutaBackup & "\PRODUCC.MDB"
  DataAct.Refresh
  DataOld.DatabaseName = RutaBackup & "\PRODUCC.MDB"
  DataOld.Refresh
  Respaldos.Caption = "Procesando Facturas..."
  sSQL = "SELECT * FROM Clientes "
  SelectData DataQuery, sSQL, False
  Respaldos.Caption = "Procesando Facturas..........."
  sSQL = "DELETE * FROM Facturas "
  sSQL = sSQL & "WHERE NOT Fecha BETWEEN #" & BuscarFecha(MBoxFechaI.Text) & "# "
  sSQL = sSQL & "and #" & BuscarFecha(MBoxFechaF.Text) & "# "
  DeleteData DataQuery, sSQL
  Respaldos.Caption = "Procesando Detalle de Factura..........."
  sSQL = "DELETE * FROM Detalle_Factura "
  sSQL = sSQL & "WHERE NOT Fecha BETWEEN #" & BuscarFecha(MBoxFechaI.Text) & "# "
  sSQL = sSQL & "and #" & BuscarFecha(MBoxFechaF.Text) & "# "
  DeleteData DataQuery, sSQL
  Respaldos.Caption = "Procesando Contratos..."
  sSQL = "DELETE * FROM Contratos "
  sSQL = sSQL & "WHERE NOT Desde BETWEEN #" & BuscarFecha(MBoxFechaI.Text) & "# "
  sSQL = sSQL & "and #" & BuscarFecha(MBoxFechaF.Text) & "# "
  DeleteData DataQuery, sSQL
  sSQL = "DELETE * FROM Parte_Contratos "
  sSQL = sSQL & "WHERE NOT Fecha BETWEEN #" & BuscarFecha(MBoxFechaI.Text) & "# "
  sSQL = sSQL & "and #" & BuscarFecha(MBoxFechaF.Text) & "# "
  DeleteData DataQuery, sSQL
  Respaldos.Caption = "Procesando Diario de Caja..."
  sSQL = "DELETE * FROM Diario_Caja "
  sSQL = sSQL & "WHERE NOT Fecha BETWEEN #" & BuscarFecha(MBoxFechaI.Text) & "# "
  sSQL = sSQL & "and #" & BuscarFecha(MBoxFechaF.Text) & "# "
  DeleteData DataQuery, sSQL
 'Clientes
  Respaldos.Caption = "Procesando Clientes..."
  sSQL = "SELECT Codigo FROM Clientes "
  sSQL = sSQL & "GROUP BY Codigo "
  No_Desde = 0: No_Hasta = 0: Contador = 0
  SelectData DataQuery, sSQL, False
  With DataQuery.Recordset
    Do While Not .EOF
       Codigo = .Fields("Codigo"): Contador = Contador + 1
       Respaldos.Caption = "Registro de Clientes: " & Contador & "/" & .RecordCount & String(Val(Mid(Str(Contador), Len(Str(Contador)), 1)), "||")
       sSQL = "SELECT * FROM Diario_Caja "
       sSQL = sSQL & "WHERE Codigo_C = '" & Codigo & "' "
       SelectData DataOld, sSQL
       If DataOld.Recordset.RecordCount <= 0 Then
          sSQL = "DELETE * FROM Clientes "
          sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
          DeleteData DataOld, sSQL
          sSQL = "DELETE * FROM Clientes_Aux "
          sSQL = sSQL & "WHERE Codigo = '" & Codigo & "' "
          DeleteData DataOld, sSQL
       End If
      .MoveNext
    Loop
  End With
 'GenerarArchivoPlano Respaldos, DataQuery, "B_Envios.txt"
  DataQuery.Database.Close
  DataAct.Database.Close
  DataOld.Database.Close
  Respaldos.Caption = "Compactando las bases en Access..."
  For Ind = 0 To File1.ListCount - 1
      Cadena = Dir1.Path & "\" & File1.List(Ind)
      CompactarDatabase Cadena, False
  Next Ind
  Respaldos.Caption = "Compactando las bases en PKZIP..."
  Shell RutaBackup & "\BFACTURA.BAT " & TextUnidad.Text & ": ", vbMaximizedFocus
  RatonNormal
  Respaldos.Caption = "RESPALDOS DE BASES"
  Unload Respaldos
End Sub

Private Sub Command2_Click()
  Unload Respaldos
End Sub

Private Sub Command3_Click()
'Envios cancelados de sucursales
  RatonReloj
  FechaValida MBoxFechaI
  Respaldos.Caption = "Copiando bases requeridas"
  Dir1.Path = RutaEmpresa
  File1.filename = Dir1.Path & "\ENV*.MDB"
  For Ind = 0 To File1.ListCount - 1
      RutaOrigen = Dir1.Path & "\" & File1.List(Ind)
      RutaDestino = RutaBackup & "\" & File1.List(Ind)
      Respaldos.Caption = "Copiando la base: " & RutaDestino
      FileCopy RutaOrigen, RutaDestino
  Next Ind
  RutaOrigen = RutaSistema & "\Empresas.mdb"
  RutaDestino = RutaBackup & "\Empresas.mdb"
  FileCopy RutaOrigen, RutaDestino
  RutaOrigen = RutaSistema & "\pkzip.exe"
  RutaDestino = RutaBackup & "\pkzip.exe"
  FileCopy RutaOrigen, RutaDestino
  RutaOrigen = RutaSistema & "\pkunzip.exe"
  RutaDestino = RutaBackup & "\pkunzip.exe"
  FileCopy RutaOrigen, RutaDestino
  Dir1.Path = RutaBackup
  File1.filename = Dir1.Path & "\*.MDB"
  DataQuery.DatabaseName = RutaBackup & "\ENVIOS.MDB"
  DataQuery.Refresh
  DataOld.DatabaseName = RutaBackup & "\ENVIOS.MDB"
  DataOld.Refresh
  Respaldos.Caption = "Procesando Llamadas..."
  sSQL = "DELETE * FROM Resumen_Llamadas "
  DeleteData DataQuery, sSQL
  MiFecha = CLongFecha(CFechaLong(MBoxFechaI.Text) - 6)
  sSQL = "INSERT INTO Resumen_Llamadas " _
       & "SELECT Fecha,Envio_No,Llamadas " _
       & "FROM Correos " _
       & "WHERE Fecha BETWEEN #" & BuscarFecha(MiFecha) & "# " _
       & "and #" & BuscarFecha(MBoxFechaI.Text) & "# " _
       & "AND Llamadas <> '" & Ninguno & "' "
  InsertData DataQuery, sSQL
  Respaldos.Caption = "Procesando Correos..."
  sSQL = "DELETE * FROM Correos "
  sSQL = sSQL & "WHERE T <> 'C' "
  DeleteData DataQuery, sSQL
  sSQL = "DELETE * FROM Correos "
  sSQL = sSQL & "WHERE Fecha_P <> #" & BuscarFecha(MBoxFechaI.Text) & "# "
  DeleteData DataQuery, sSQL
  Respaldos.Caption = "Procesando Flujo de Caja..."
  sSQL = "DELETE * FROM Flujo_Caja "
  sSQL = sSQL & "WHERE Fecha <> #" & BuscarFecha(MBoxFechaI.Text) & "# "
  DeleteData DataQuery, sSQL
 'Beneficiarios
  No_Desde = 0: No_Hasta = 0: Contador = 0
  Respaldos.Caption = "Procesando Beneficiarios..."
  sSQL = "SELECT Cod_B FROM Correos "
  sSQL = sSQL & "GROUP BY Cod_B "
  SelectData DataQuery, sSQL, False
  With DataQuery.Recordset
    Do While Not .EOF
       No_Hasta = .Fields("Cod_B"): Contador = Contador + 1
       Respaldos.Caption = "Registro de Beneficiario: " & Contador & "/" & .RecordCount & String(Val(Mid(Str(Contador), Len(Str(Contador)), 1)), "|")
       sSQL = "DELETE * FROM Beneficiarios "
       sSQL = sSQL & "WHERE " & No_Desde & " < Codigo_B "
       sSQL = sSQL & "AND Codigo_B < " & No_Hasta & " "
       DeleteData DataOld, sSQL
      .MoveNext
      No_Desde = No_Hasta
    Loop
    sSQL = "DELETE * FROM Beneficiarios "
    sSQL = sSQL & "WHERE Codigo_B > " & No_Hasta & " "
    DeleteData DataOld, sSQL
  End With
 'Remitentes
  Respaldos.Caption = "Procesando Remitentes..."
  sSQL = "SELECT Cod_R FROM Correos "
  sSQL = sSQL & "GROUP BY Cod_R "
  No_Desde = 0: No_Hasta = 0: Contador = 0
  SelectData DataQuery, sSQL, False
  With DataQuery.Recordset
    Do While Not .EOF
       No_Hasta = .Fields("Cod_R"): Contador = Contador + 1
       Respaldos.Caption = "Registro de Remitente: " & Contador & "/" & .RecordCount & String(Val(Mid(Str(Contador), Len(Str(Contador)), 1)), "|")
       sSQL = "DELETE * FROM Remitentes "
       sSQL = sSQL & "WHERE " & No_Desde & " < Codigo_R "
       sSQL = sSQL & "AND Codigo_R < " & No_Hasta & " "
       DeleteData DataOld, sSQL
      .MoveNext
      No_Desde = No_Hasta
    Loop
    sSQL = "DELETE * FROM Remitentes "
    sSQL = sSQL & "WHERE Codigo_R > " & No_Hasta & " "
    DeleteData DataOld, sSQL
  End With
  DataQuery.Database.Close
  DataOld.Database.Close
  Respaldos.Caption = "Compactando las bases..."
  For Ind = 0 To File1.ListCount - 1
      Cadena = Dir1.Path & "\" & File1.List(Ind)
      CompactarDatabase Cadena, False
  Next Ind
  Shell RutaBackup & "\BENVIOS.BAT SYSBASES ", vbMaximizedFocus
  RatonNormal
  Respaldos.Caption = "RESPALDOS DE BASES"
  Unload Respaldos
End Sub

Private Sub Command4_Click()
'actualizar cancelados de sucursales
RatonReloj
If Una_Vez Then
  Respaldos.Caption = "Espere, descomprimiendo bases de datos..."
  'MsgBox Cadena
  RatonNormal
  Contador = 1
  CodSucursal = Ninguno
  DataAux.DatabaseName = RutaBackup & "\ENVIOS.MDB"
  DataAux.Refresh
  sSQL = "SELECT T,Fecha_P,Sucursal FROM Correos "
  sSQL = sSQL & "WHERE T = 'C' "
  SelectData DataAux, sSQL, False
  If DataAux.Recordset.RecordCount > 0 Then
     MiFecha = DataAux.Recordset.Fields("Fecha_P")
     CodSucursal = DataAux.Recordset.Fields("Sucursal")
  End If
  DataQuery.DatabaseName = RutaEmpresa & "\ENVIOS.MDB"
  DataQuery.Refresh
  Cadena = ""
  sSQL = "SELECT * FROM Sucursales "
  sSQL = sSQL & "WHERE Codigo = '" & CodSucursal & "' "
  SelectData DataQuery, sSQL, False
  With DataQuery.Recordset
   If .RecordCount > 0 Then
       Cadena = "Sucursal: " & .Fields("Ciudad") & Chr(13) & Chr(13) _
              & "Fecha   : " & FechaStrgDias(MiFecha)
   End If
  End With
  DataQuery.Database.Close
  DataAux.Database.Close
  If Cadena <> "" Then
     Titulo = "Proceso de Actualizacion"
     Mensajes = Cadena & Chr(13) & "Quiere actualizar"
     If BoxMensaje = 6 Then
        DataOld.DatabaseName = RutaBackup & "\ENVIOS.MDB"
        DataOld.Refresh
        DataAux.DatabaseName = RutaBackup & "\ENVIOS.MDB"
        DataAux.Refresh
        DataAct.DatabaseName = RutaEmpresa & "\ENVIOS.MDB"
        DataAct.Refresh
        Respaldos.Caption = "Actualizando Llamadas"
        sSQL = "SELECT * FROM Resumen_Llamadas "
        sSQL = sSQL & "ORDER BY Envio_No "
        SelectData DataOld, sSQL
        RatonReloj
        With DataOld.Recordset
         If .RecordCount > 0 Then
             Do While Not .EOF
                Codigo = .Fields("Envio_No")
                Cadena = .Fields("Llamadas")
                If Cadena = "" Then Cadena = Ninguno
                Respaldos.Caption = "Envio No. " & Codigo
                
                sSQL = "UPDATE Correos "
                sSQL = sSQL & "SET Llamadas = '" & Cadena & "' "
                sSQL = sSQL & "WHERE Envio_No = '" & Codigo & "' "
                UpdateData DataAct, sSQL
               .MoveNext
             Loop
         End If
        End With
        sSQL = "DELETE * FROM Flujo_Caja "
        sSQL = sSQL & "WHERE Fecha = #" & BuscarFecha(MiFecha) & "# "
        sSQL = sSQL & "AND Sucursal = '" & Format(CodSucursal, "000") & "' "
        DeleteData DataAct, sSQL
        
        sSQL = "SELECT * FROM Correos "
        sSQL = sSQL & "WHERE T = 'C' "
        SelectData DataOld, sSQL, False
        sSQL = "SELECT * FROM Correos "
        SelectData DataAct, sSQL, False
        If DataOld.Recordset.RecordCount > 0 Then
           RatonReloj
           TotalReg = DataOld.Recordset.RecordCount
           Do While Not DataOld.Recordset.EOF
              Codigo2 = DataOld.Recordset.Fields("Envio_No")
              MiFecha = DataOld.Recordset.Fields("Fecha_P")
              Cod_Benef = DataOld.Recordset.Fields("Cod_B")
              Cod_Remit = DataOld.Recordset.Fields("Cod_R")
              Respaldos.Caption = Contador & "/" & TotalReg & " Envio No. " & Codigo2
              sSQL = "DELETE * FROM Correos "
              sSQL = sSQL & "WHERE Envio_No = '" & Codigo2 & "' "
              DeleteData DataAct, sSQL
              sSQL = "SELECT * FROM Correos "
              sSQL = sSQL & "WHERE Envio_No = '" & Codigo2 & "' "
              SelectData DataAct, sSQL, False
              DataAct.Recordset.AddNew
              DataAct.Recordset.Fields("T") = Cancelado
              DataAct.Recordset.Fields("CI") = DataOld.Recordset.Fields("CI")
              DataAct.Recordset.Fields("IP") = DataOld.Recordset.Fields("IP")
              DataAct.Recordset.Fields("ME") = DataOld.Recordset.Fields("ME")
              DataAct.Recordset.Fields("Fecha") = DataOld.Recordset.Fields("Fecha")
              DataAct.Recordset.Fields("Fecha_P") = DataOld.Recordset.Fields("Fecha_P")
              DataAct.Recordset.Fields("Recibo_No") = DataOld.Recordset.Fields("Recibo_No")
              DataAct.Recordset.Fields("Envio_No") = DataOld.Recordset.Fields("Envio_No")
              DataAct.Recordset.Fields("Giro_No") = DataOld.Recordset.Fields("Giro_No")
              DataAct.Recordset.Fields("Cod_R") = DataOld.Recordset.Fields("Cod_R")
              DataAct.Recordset.Fields("Cod_B") = DataOld.Recordset.Fields("Cod_B")
              DataAct.Recordset.Fields("Cod_C") = DataOld.Recordset.Fields("Cod_C")
              DataAct.Recordset.Fields("Cantidad") = DataOld.Recordset.Fields("Cantidad")
              DataAct.Recordset.Fields("Porc_C") = DataOld.Recordset.Fields("Porc_C")
              DataAct.Recordset.Fields("TOTAL") = DataOld.Recordset.Fields("TOTAL")
              DataAct.Recordset.Fields("PAGADO") = DataOld.Recordset.Fields("PAGADO")
              DataAct.Recordset.Fields("SALDO") = DataOld.Recordset.Fields("SALDO")
              DataAct.Recordset.Fields("Mensaje") = DataOld.Recordset.Fields("Mensaje")
              DataAct.Recordset.Fields("Cotizacion") = DataOld.Recordset.Fields("Cotizacion")
              DataAct.Recordset.Fields("Usuario") = DataOld.Recordset.Fields("Usuario")
              DataAct.Recordset.Fields("Sucursal") = DataOld.Recordset.Fields("Sucursal")
              DataAct.Recordset.Fields("Llamadas") = DataOld.Recordset.Fields("Llamadas")
              DataAct.Recordset.Update
             'Actualizar Beneficiarios
              sSQL = "SELECT * FROM Beneficiarios "
              sSQL = sSQL & "WHERE Codigo_B = " & Cod_Benef & " "
              SelectData DataAux, sSQL, False
              Codigo = DataAux.Recordset.Fields("CI_RUC")
              Codigo1 = DataAux.Recordset.Fields("Telefono")
              Codigo2 = DataAux.Recordset.Fields("Direccion")
              Respaldos.Caption = FechaStrgDias(MiFecha) & " => " & CodSucursal & ". Actualizando Beneficiario: " & Cod_Benef & "..."
              sSQL = "SELECT * FROM Beneficiarios "
              SelectData DataAct, sSQL, False
              sSQL = "UPDATE Beneficiarios "
              sSQL = sSQL & "SET CI_RUC = '" & Codigo & "', "
              sSQL = sSQL & "Telefono = '" & Codigo1 & "', "
              sSQL = sSQL & "Direccion = '" & Codigo2 & "' "
              sSQL = sSQL & "WHERE Codigo_B = " & Cod_Benef & " "
              UpdateData DataAct, sSQL
              Contador = Contador + 1
              DataOld.Recordset.MoveNext
           Loop
           RatonNormal
        End If
        Respaldos.Caption = "Actualizando Flujo de Caja..."
        sSQL = "SELECT * FROM Flujo_Caja "
        sSQL = sSQL & "ORDER BY Sucursal "
        SelectData DataOld, sSQL, False
        If DataOld.Recordset.RecordCount > 0 Then
           RatonReloj
           Codigo = DataOld.Recordset.Fields("Sucursal")
           sSQL = "SELECT * FROM Flujo_Caja "
           SelectData DataAct, sSQL, False
           sSQL = "DELETE * FROM Flujo_Caja "
           sSQL = sSQL & "WHERE Sucursal = '" & Codigo & "' "
           DeleteData DataAct, sSQL
           Do While Not DataOld.Recordset.EOF
              If Codigo <> DataOld.Recordset.Fields("Sucursal") Then
                 Codigo = DataOld.Recordset.Fields("Sucursal")
                 sSQL = "DELETE * FROM Flujo_Caja "
                 sSQL = sSQL & "WHERE Sucursal = '" & Codigo & "' "
                 DeleteData DataAct, sSQL
              End If
              DataAct.Recordset.AddNew
              DataAct.Recordset.Fields("Fecha") = DataOld.Recordset.Fields("Fecha")
             'DataAct.Recordset.Fields("Concepto") = DataOld.Recordset.Fields("Concepto")
              DataAct.Recordset.Fields("Envio_No") = DataOld.Recordset.Fields("Envio_No")
              DataAct.Recordset.Fields("Ingreso_ME") = DataOld.Recordset.Fields("Ingreso_ME")
              DataAct.Recordset.Fields("Egreso_ME") = DataOld.Recordset.Fields("Egreso_ME")
              DataAct.Recordset.Fields("Ingreso_MN") = DataOld.Recordset.Fields("Ingreso_MN")
              DataAct.Recordset.Fields("Egreso_MN") = DataOld.Recordset.Fields("Egreso_MN")
              DataAct.Recordset.Fields("Usuario") = DataOld.Recordset.Fields("Usuario")
              DataAct.Recordset.Fields("C") = DataOld.Recordset.Fields("C")
              DataAct.Recordset.Fields("Sucursal") = DataOld.Recordset.Fields("Sucursal")
              DataAct.Recordset.Update
              DataOld.Recordset.MoveNext
           Loop
           RatonNormal
        End If
        DataOld.Database.Close
        DataAct.Database.Close
        DataAux.Database.Close
     End If
  End If
  Else
    MsgBox "primero descomprima "
    RatonNormal
  End If
  Unload Respaldos
End Sub

Private Sub Command5_Click()
'Actualizar pendientes sucursales
Dim YaExiste As Boolean
  YaExiste = False
  RatonReloj
  If Una_Vez Then
     Respaldos.Caption = "Espere, descomprimiendo bases de datos..."
     RatonNormal
     DataOld.DatabaseName = RutaBackup & "\PRODUCC.MDB"
     DataOld.Refresh
     DataQuery.DatabaseName = RutaBackup & "\PRODUCC.MDB"
     DataQuery.Refresh
     DataAct.DatabaseName = RutaEmpresa & "\PRODUCC.MDB"
     DataAct.Refresh
     DataEmpresas.DatabaseName = RutaSistema & "\EMPRESAS.MDB"
     DataEmpresas.Refresh
     Contador = 1
     'DateValue
     sSQL = "SELECT * FROM Facturas "
     SelectData DataAct, sSQL
     Mensajes = "RESULTADOS DE ACTUALIZACION:" & Chr(13) & Chr(13)
     sSQL = "SELECT * FROM Empresas "
     sSQL = sSQL & "ORDER BY Item "
     SelectData DataEmpresas, sSQL
     With DataEmpresas.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             CantCtas = .Fields("Item")
             Mensaje = .Fields("Empresa")
             Label1.Caption = "Actualizando: " & CantCtas & " - " & Mensaje
             Contador = 0: Rotos = 0
             sSQL = "INSERT INTO Facturas "
             sSQL = sSQL & "IN '" & RutaEmpresa & "\PRODUCC.MDB' "
             sSQL = sSQL & "SELECT * FROM Facturas "
             sSQL = sSQL & "WHERE Factura <> Factura "
             InsertData DataOld, sSQL
            .MoveNext
          Loop
      End If
     End With
  Else
     MsgBox "primero descomprima "
     RatonNormal
  End If
  Label1.Caption = Empresa
  RatonNormal
  Unload Respaldos
End Sub

Private Sub Command6_Click()
'Primero descomprima para actualizar
  RatonReloj
  Shell RutaBackup & "\RFACTURA.BAT " & TextUnidad.Text & ": ", vbMaximizedFocus
  Respaldos.Caption = "RESPALDOS DE BASES"
  RatonNormal
  Una_Vez = True
End Sub

Private Sub Command7_Click()
End Sub

Private Sub Form_Activate()
  Una_Vez = False
  RutaBackup = Left(CurDir$, 2) & "\SYSBASES"
  Dir1.Path = RutaEmpresa
  File1.filename = Dir1.Path & "\*.*"
  Respaldos.Caption = "RESPALDOS DE BASES"
  Label1.Caption = Empresa
  RatonNormal
  MBoxFechaI.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm Respaldos
End Sub

Private Sub MBoxFechaI_GotFocus()
  MarcarTexto MBoxFechaI
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub

Private Sub MBoxFechaF_GotFocus()
  MarcarTexto MBoxFechaF
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub TextUnidad_GotFocus()
  MarcarTexto TextUnidad
End Sub

Private Sub TextUnidad_LostFocus()
  TextoValido TextUnidad, True
End Sub
