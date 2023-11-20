VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form Respaldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"Respaldo.frx":0000
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "Respaldo.frx":0092
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
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
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "A"
      Top             =   735
      Width           =   1275
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
      Height          =   960
      Left            =   105
      Picture         =   "Respaldo.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1155
      Width           =   2430
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Descomprimir Bases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      Picture         =   "Respaldo.frx":0916
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2205
      Width           =   2430
   End
   Begin MSMask.MaskEdBox MBoxFechaF 
      Height          =   330
      Left            =   1260
      TabIndex        =   3
      Top             =   420
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
   Begin VB.CommandButton Command3 
      Caption         =   "&Visualizar Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      Picture         =   "Respaldo.frx":0D58
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3255
      Width           =   2430
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Actualizar de Agencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      Picture         =   "Respaldo.frx":119A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4305
      Width           =   2430
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
      Height          =   960
      Left            =   105
      Picture         =   "Respaldo.frx":15DC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5355
      Width           =   2430
   End
   Begin VB.Data DataOld 
      Caption         =   "Old"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3045
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Data DataAux 
      Caption         =   "Aux"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3045
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1470
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Data DataAct 
      Caption         =   "Act"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3045
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1155
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   3045
      TabIndex        =   12
      Top             =   2625
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3045
      TabIndex        =   11
      Top             =   2310
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
      Left            =   3045
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1785
      Visible         =   0   'False
      Width           =   1800
   End
   Begin MSMask.MaskEdBox MBoxFechaI 
      Height          =   330
      Left            =   1260
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
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   2625
      TabIndex        =   15
      Top             =   6510
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
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
      TabIndex        =   4
      Top             =   735
      Width           =   1170
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Fina."
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
      TabIndex        =   2
      Top             =   420
      Width           =   1170
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
      Left            =   2625
      TabIndex        =   6
      Top             =   105
      Width           =   8310
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha Inic."
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
      Top             =   105
      Width           =   1170
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
  TextoValido TextUnidad, , True
  FechaValida MBoxFechaI, False
  FechaValida MBoxFechaF, False
  FechaIni = BuscarFecha(MBoxFechaI.Text)
  FechaFin = BuscarFecha(MBoxFechaF.Text)
  Respaldos.Caption = "Copiando bases requeridas"
  Dir1.Path = RutaEmpresa
  File1.FileName = Dir1.Path & "\CAJACR*.MDB"
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
  File1.FileName = Dir1.Path & "\*.MDB"
  DataQuery.DatabaseName = RutaBackup & "\CAJACRED.MDB"
  DataQuery.Refresh
  DataAct.DatabaseName = RutaBackup & "\CAJACRED.MDB"
  DataAct.Refresh
  DataOld.DatabaseName = RutaBackup & "\CAJACRED.MDB"
  DataOld.Refresh
  
  sSQL = "SELECT * FROM Cuentas "
  SelectData DataQuery, sSQL, False
  
  Respaldos.Caption = "Procesando Cuentas..."
  sSQL = "DELETE * FROM Cuentas "
  sSQL = sSQL & "WHERE Fecha Not Between #" & FechaIni & "# and #" & FechaFin & "# "
  DeleteData DataQuery, sSQL
  
  Respaldos.Caption = "Procesando Bancos..."
  sSQL = "DELETE * FROM Bancos "
  sSQL = sSQL & "WHERE Fecha Not Between #" & FechaIni & "# and #" & FechaFin & "# "
  DeleteData DataQuery, sSQL
  
  sSQL = "UPDATE Bancos "
  sSQL = sSQL & "SET Cta = '.' "
  UpdateData DataQuery, sSQL
  
  Respaldos.Caption = "Procesando Bloqueos......."
  sSQL = "DELETE * FROM Bloqueos "
  sSQL = sSQL & "WHERE Fecha Not Between #" & FechaIni & "# and #" & FechaFin & "# "
  DeleteData DataQuery, sSQL
  
  Respaldos.Caption = "Procesando Intereses..."
  sSQL = "DELETE * FROM Intereses "
  sSQL = sSQL & "WHERE Fecha Not Between #" & FechaIni & "# and #" & FechaFin & "# "
  DeleteData DataQuery, sSQL
  
  Respaldos.Caption = "Procesando Prestamos..."
  sSQL = "DELETE * FROM Prestamos "
  sSQL = sSQL & "WHERE Fecha Not Between #" & FechaIni & "# and #" & FechaFin & "# "
  DeleteData DataQuery, sSQL
  
  sSQL = "DELETE * FROM Prestamos "
  sSQL = sSQL & "WHERE T = 'N' "
  DeleteData DataQuery, sSQL
  
  Respaldos.Caption = "Procesando Saldo de Cajas o Libretas..."
  sSQL = "DELETE * FROM Saldo_Caja_Libreta "
  sSQL = sSQL & "WHERE Fecha Not Between #" & FechaIni & "# and #" & FechaFin & "# "
  DeleteData DataQuery, sSQL
  
  Respaldos.Caption = "Procesando Saldo de Libretas..."
  sSQL = "DELETE * FROM Saldo_Libretas "
  sSQL = sSQL & "WHERE Fecha Not Between #" & FechaIni & "# and #" & FechaFin & "# "
  DeleteData DataQuery, sSQL
  
  Respaldos.Caption = "Procesando Transacciones de Cajas..."
  sSQL = "DELETE * FROM Trans_Cajas "
  sSQL = sSQL & "WHERE Fecha Not Between #" & FechaIni & "# and #" & FechaFin & "# "
  DeleteData DataQuery, sSQL
  
  Respaldos.Caption = "Procesando Transacciones de Libretas..."
  sSQL = "DELETE * FROM Trans_Libretas "
  sSQL = sSQL & "WHERE Fecha Not Between #" & FechaIni & "# and #" & FechaFin & "# "
  DeleteData DataQuery, sSQL
  Contador = 0
  Respaldos.Caption = "Procesando Transacciones de Prestamos..."
  sSQL = "SELECT Cuenta_No,Credito_No "
  sSQL = sSQL & "FROM Trans_Prestamos "
  sSQL = sSQL & "GROUP BY Cuenta_No,Credito_No "
  SelectData DataOld, sSQL
  With DataOld.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Cuenta_No = .Fields("Cuenta_No")
          Numero = .Fields("Credito_No")
          Suma_ME = (Contador * 100) / (.RecordCount - 1)
          Respaldos.Caption = Round(Suma_ME) & "% Procesando Transacciones de Prestamos..."
          sSQL = "SELECT * FROM Prestamos "
          sSQL = sSQL & "WHERE Cuenta_No = '" & Cuenta_No & "' "
          sSQL = sSQL & "AND Credito_No = " & Numero & " "
          SelectData DataAct, sSQL
          If DataAct.Recordset.RecordCount <= 0 Then
             sSQL = "DELETE * FROM Trans_Prestamos "
             sSQL = sSQL & "WHERE Fecha_C Not Between #" & FechaIni & "# and #" & FechaFin & "# "
             sSQL = sSQL & "AND Cuenta_No = '" & Cuenta_No & "' "
             sSQL = sSQL & "AND Credito_No = " & Numero & " "
             DeleteData DataQuery, sSQL
          End If
          Contador = Contador + 1
         .MoveNext
       Loop
   End If
  End With
  
  DataQuery.Database.Close
  DataAct.Database.Close
  DataOld.Database.Close
  
  Respaldos.Caption = "Compactando las bases en Access..."
  For Ind = 0 To File1.ListCount - 1
      Cadena = Dir1.Path & "\" & File1.List(Ind)
      CompactarDatabase Cadena, False
  Next Ind
  
  Respaldos.Caption = "Compactando las bases en PKZIP..."
  Shell RutaBackup & "\BCAJACRE.BAT " & UCase(TextUnidad.Text) & ": ", vbMaximizedFocus
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
  Dir1.Path = RutaEmpresa
  File1.FileName = Dir1.Path & "\CAJACR*.MDB"
 'Abrir Bases de Datos
  DataQuery.DatabaseName = RutaBackup & "\CAJACRED.MDB"
  DataQuery.Refresh
  DataOld.DatabaseName = RutaBackup & "\CAJACRED.MDB"
  DataOld.Refresh
  Respaldos.Caption = "Procesando Depositos/Retiros de Agencias..."
 'Procesar Dep/Ret de Agencias
  No_Desde = 0: No_Hasta = 0: Contador = 0
  Respaldos.Caption = "Procesando Beneficiarios..."
  sSQL = "SELECT Fecha,Cuenta_No,ME,TP.TP,TL.Cheque,Debitos,Creditos,Item "
  sSQL = sSQL & "FROM Tipo_Proceso As TP,Trans_Libretas As TL "
  sSQL = sSQL & "WHERE TL.TP = TP.TP "
  sSQL = sSQL & "ORDER BY Fecha,Cuenta_No,TP.TP "
  SelectDBGrid DBGQuery, DataQuery, sSQL
  Respaldos.Caption = "Compactando las bases..."
  RatonNormal
  'Unload Respaldos
End Sub

Private Sub Command4_Click()
'Actualizacion de sucursales
  ProgBarra.Max = 100
  ProgBarra.Min = 0
  Contador = 0
  ProgBarra.Value = 0
  DBGQuery.Visible = False
  FechaIni = "": FechaFin = ""
  NumItem = NumEmpresa
  DataAux.DatabaseName = RutaBackup & "\CAJACRED.MDB"
  DataAux.Refresh
  sSQL = "SELECT TP,Cuenta_No,Fecha,Item "
  sSQL = sSQL & "FROM Trans_Libretas "
  sSQL = sSQL & "ORDER BY Fecha "
  SelectData DataAux, sSQL, False
  If DataAux.Recordset.RecordCount > 0 Then
     DataAux.Recordset.MoveFirst
     NumItem = DataAux.Recordset.Fields("Item")
     FechaIni = DataAux.Recordset.Fields("Fecha")
     DataAux.Recordset.MoveLast
     FechaFin = DataAux.Recordset.Fields("Fecha")
  End If
  DataQuery.DatabaseName = RutaBackup & "\EMPRESAS.MDB"
  DataQuery.Refresh
  Cadena = ""
  sSQL = "SELECT * FROM Empresas "
  sSQL = sSQL & "WHERE Item = " & NumItem & " "
  SelectData DataQuery, sSQL, False
  With DataQuery.Recordset
   If .RecordCount > 0 Then
       Cadena = "Empresa: " & .Fields("Empresa") & Chr(13) & Chr(13) _
              & "Rango de fecha: " & FechaIni & " - " & FechaFin
   End If
  End With
  DataQuery.Database.Close
  DataAux.Database.Close
  DataAux.DatabaseName = RutaBackup & "\CAJACRED.MDB"
  DataAux.Refresh
  DataOld.DatabaseName = RutaBackup & "\CAJACRED.MDB"
  DataOld.Refresh
  DataAct.DatabaseName = RutaEmpresa & "\CAJACRED.MDB"
  DataAct.Refresh
  If Cadena <> "" And FechaIni <> "" And FechaFin <> "" Then
     Titulo = "Proceso de Actualizacion"
     Mensajes = Cadena & Chr(13) & "Quiere Actualizar las Bases"
     If BoxMensaje = 6 Then
        RatonReloj
        FechaIni = BuscarFecha(FechaIni)
        FechaFin = BuscarFecha(FechaFin)
        ProgBarra.Value = 4
       'If OpcCoop Then sSQL = sSQL & "AND Item = " & NumItem & " "
       'If OpcCoop = False And CheqCoop.Value = 1 Then sSQL = sSQL & "AND Item = " & NumItem & " "
        sSQL = "DELETE * FROM Cuentas "
        sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
        DeleteData DataAct, sSQL
        ProgBarra.Value = 8
        sSQL = "DELETE * FROM Bancos "
        sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
        DeleteData DataAct, sSQL
        ProgBarra.Value = 12
        sSQL = "DELETE * FROM Bloqueos "
        sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
        DeleteData DataAct, sSQL
        ProgBarra.Value = 16
        sSQL = "DELETE * FROM Intereses "
        sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
        DeleteData DataAct, sSQL
        ProgBarra.Value = 20
        sSQL = "DELETE * FROM Prestamos "
        sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
        DeleteData DataAct, sSQL
        ProgBarra.Value = 24
        sSQL = "DELETE * FROM Saldo_Caja_Libreta "
        sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
        DeleteData DataAct, sSQL
        ProgBarra.Value = 27
        sSQL = "DELETE * FROM Saldo_Libretas "
        sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
        DeleteData DataAct, sSQL
        ProgBarra.Value = 29
        sSQL = "DELETE * FROM Trans_Cajas "
        sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
        DeleteData DataAct, sSQL
        ProgBarra.Value = 32
        sSQL = "DELETE * FROM Trans_Libretas "
        sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# "
        DeleteData DataAct, sSQL
        ProgBarra.Value = 36
        sSQL = "SELECT * FROM Trans_Prestamos "
        sSQL = sSQL & "ORDER BY Fecha "
        SelectData DataOld, sSQL
        If DataOld.Recordset.RecordCount > 0 Then
           Do While Not DataOld.Recordset.EOF
              FechaIni = BuscarFecha(DataOld.Recordset.Fields("Fecha"))
              Cuenta_No = DataOld.Recordset.Fields("Cuenta_No")
              Numero = DataOld.Recordset.Fields("Credito_No")
              sSQL = "DELETE * FROM Trans_Prestamos "
              sSQL = sSQL & "WHERE Fecha = #" & FechaIni & "# "
              sSQL = sSQL & "AND Cuenta_No = '" & Cuenta_No & "' "
              sSQL = sSQL & "AND Credito_No = " & Numero & " "
              DeleteData DataAct, sSQL
              DataOld.Recordset.MoveNext
           Loop
        End If
        sSQL = "DELETE * FROM Conyugue "
        DeleteData DataAct, sSQL
        sSQL = "DELETE * FROM Garantes "
        DeleteData DataAct, sSQL
       'Insertamos las transacciones de Caja/Creditos
        ProgBarra.Value = 40
        sSQL = "INSERT INTO Cuentas "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Cuentas "
        InsertData DataOld, sSQL
        ProgBarra.Value = 44
        sSQL = "INSERT INTO Bancos "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Bancos "
        InsertData DataOld, sSQL
        ProgBarra.Value = 48
        sSQL = "INSERT INTO Bloqueos "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Bloqueos "
        InsertData DataOld, sSQL
        ProgBarra.Value = 52
        sSQL = "INSERT INTO Intereses "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Intereses "
        InsertData DataOld, sSQL
        ProgBarra.Value = 56
        sSQL = "INSERT INTO Prestamos "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Prestamos "
        InsertData DataOld, sSQL
        ProgBarra.Value = 60
        sSQL = "INSERT INTO Saldo_Caja_Libreta "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Saldo_Caja_Libreta "
        InsertData DataOld, sSQL
        ProgBarra.Value = 62
        sSQL = "INSERT INTO Saldo_Libretas "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Saldo_Libretas "
        InsertData DataOld, sSQL
        ProgBarra.Value = 64
        sSQL = "INSERT INTO Trans_Cajas "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Trans_Cajas "
        InsertData DataOld, sSQL
        ProgBarra.Value = 68
        sSQL = "INSERT INTO Trans_Libretas "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Trans_Libretas "
        InsertData DataOld, sSQL
        ProgBarra.Value = 72
        sSQL = "INSERT INTO Trans_Prestamos "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Trans_Prestamos "
        InsertData DataOld, sSQL
        ProgBarra.Value = 76
        sSQL = "INSERT INTO Conyugue "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Conyugue "
        InsertData DataOld, sSQL
        ProgBarra.Value = 80
        sSQL = "INSERT INTO Garantes "
        sSQL = sSQL & "IN '" & RutaEmpresa & "\CAJACRED.MDB' "
        sSQL = sSQL & "SELECT * FROM Garantes "
        InsertData DataOld, sSQL
     End If
  End If
  ProgBarra.Value = 84
  DataOld.Database.Close
  DataAct.Database.Close
  DataAux.Database.Close
  ProgBarra.Value = 100
  MsgBox "Proceso de Actualizacion terminado"
  RatonNormal
 'NumEmpresa = NumItem
  Unload Respaldos
End Sub

Private Sub Command6_Click()
'Primero descomprima para actualizar
  RatonReloj
  TextoValido TextUnidad, , True
  Respaldos.Caption = "Restaurando las bases "
  If (TextUnidad.Text & ":" = "A:") Or (TextUnidad.Text & ":" = "B:") Then
     ChDrive TextUnidad.Text & ":"
     Shell "RCAJACRE.BAT " & TextUnidad.Text & ": " & Mid(RutaSistema, 1, 2), vbMaximizedFocus
  Else
     ChDrive TextUnidad.Text & ":"
     ChDir TextUnidad.Text & ":\SYSBASES"
     Shell "RCAJACRE.BAT " & TextUnidad.Text & ": ", vbMaximizedFocus
  End If
  ChDrive Mid(RutaSistema, 1, 2)
  Respaldos.Caption = "RESPALDOS DE BASES"
  RatonNormal
  Una_Vez = True
End Sub

Private Sub Form_Activate()
  Una_Vez = False
  RutaBackup = Left(CurDir$, 2) & "\SYSBASES"
  Dir1.Path = RutaEmpresa
  File1.FileName = Dir1.Path & "\*.*"
  Respaldos.Caption = "RESPALDOS DE BASES"
  Label1.Caption = Empresa
  'If NumEmpresa = 1 Then
  '   Command1.Enabled = True
  '   Command4.Enabled = True
  '   Command5.Enabled = False
  '   Command3.Enabled = False
  'Else
  '   Command1.Enabled = False
  '   Command4.Enabled = False
  '   Command5.Enabled = True
  '   Command3.Enabled = True
  'End If
 RatonNormal
 MBoxFechaI.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm Respaldos
End Sub

Private Sub MBoxFechaI_GotFocus()
  MBoxFechaI.Text = LimpiarFechas
End Sub

Private Sub MBoxFechaI_LostFocus()
  FechaValida MBoxFechaI, False
End Sub

Private Sub MBoxFechaF_GotFocus()
  MBoxFechaF.Text = LimpiarFechas
End Sub

Private Sub MBoxFechaF_LostFocus()
  FechaValida MBoxFechaF, False
End Sub

Private Sub TextUnidad_GotFocus()
   MarcarTexto TextUnidad
End Sub

Private Sub TextUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextUnidad_LostFocus()
  TextoValido TextUnidad, , True
End Sub
