VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FBancoPichincha 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BANCO INTERNACIONAL"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "BcoPichi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "BcoPichi.frx":0442
   ScaleHeight     =   6090
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheqPend 
      BackColor       =   &H0080FFFF&
      Caption         =   "Sin Deuda Pendiente"
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
      Left            =   4305
      TabIndex        =   12
      Top             =   105
      Width           =   2325
   End
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "BcoPichi.frx":0D79
      DataSource      =   "AdoInv"
      Height          =   315
      Left            =   2415
      TabIndex        =   2
      Top             =   735
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
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
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Enviar Nomina al Banco"
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
      Left            =   6720
      Picture         =   "BcoPichi.frx":0D8E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   105
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Recibir &Abonos del Banco"
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
      Left            =   6720
      Picture         =   "BcoPichi.frx":1598
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1155
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFFF&
      Height          =   4185
      Left            =   3360
      TabIndex        =   9
      Top             =   1365
      Width           =   3270
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   5670
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   1365
      Width           =   3165
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
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
      Height          =   960
      Left            =   6720
      Picture         =   "BcoPichi.frx":1E3E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2205
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   105
      TabIndex        =   6
      Top             =   1680
      Width           =   3165
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   525
      Top             =   2205
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   525
      Top             =   2520
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
      Caption         =   "Query"
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
   Begin MSAdodcLib.Adodc AdoAct 
      Height          =   330
      Left            =   525
      Top             =   2835
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
      Caption         =   "Act"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   945
      TabIndex        =   1
      Top             =   735
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   525
      Top             =   3150
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
      Caption         =   "Act"
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
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARCHIVO:"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   1155
      Width           =   3270
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA:"
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
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &ORIGEN"
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
      Left            =   105
      TabIndex        =   4
      Top             =   1155
      Width           =   3165
   End
End
Attribute VB_Name = "FBancoPichincha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AdoStrCnnOld As String
Dim AdoStrCnn1 As String
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim XAdoStrCnn As String
Dim IJ As Long
Dim ModuloResp As String
Dim RetVal

Public Sub TipoProcesos(Opciones As String)
  NombreArchivo = ""
  Select Case Opciones
    Case "DESCUENT"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case "FACTURAS"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case "NOMINA"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case Else
         Dir1.Path = RutaBackup & "\"
         File1.FileName = Dir1.Path & "\*.*"
  End Select
  Dir1.Refresh
End Sub

Private Sub Command1_Click()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim Total_Alumnos As Long
Dim CamposFile() As Campos_Tabla
Dim EsComa As Boolean
Dim EsTab As Boolean
  TextoImprimio = ""
  sSQL = "UPDATE Facturas " _
       & "SET X = '.' " _
       & "WHERE Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND TC NOT IN ('C','P') " _
       & "AND X <> '.' " _
       & "AND T <> 'A' "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> 0 " _
       & "ORDER BY CI_RUC "
  SelectAdodc AdoAux, sSQL
  EsComa = False
  EsTab = True
  FechaValida MBFechaI
  MiFecha = BuscarFecha(MBFechaI)
  FechaTexto = MBFechaI ' FechaSistema
  CodigoInv1 = Trim(SinEspaciosIzq(DCInv))
  CodigoInv1 = Mid(CodigoInv1, Len(CodigoInv1) - 2, 3)
  DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
  RutaGeneraFile = UCase(Dir1.Path & "\" & NombreArchivo)
  TotalIngreso = 0
  Contador = 0: FileResp = 0
  ProgBarra.Value = 0
  ProgBarra.Min = 0
 'Establecemos los campos del archivo plano del Banco
  NumFile = FreeFile
  Total_Alumnos = 0
  Open RutaGeneraFile For Input As #NumFile
       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field
          If Total_Alumnos = 0 Then
          Cadena = Cod_Field
          FechaTexto = ""
          TotalReg = Len(Cadena)
          No_Desde = 1: No_Hasta = 1
          CantCampos = 0
          Do While Len(Cadena) > 0
             Do
               No_Hasta = No_Hasta + 1
               If Mid(Cadena, No_Hasta, 1) = vbTab Then EsTab = True
             Loop Until Mid(Cadena, No_Hasta, 1) = vbTab Or No_Hasta > TotalReg
            'Obtenemos la fecha de subida
'             If CantCampos = 8 Then FechaTexto = Trim(Mid(Cadena, No_Desde + 1, No_Hasta - 2)) & "/"
             CantCampos = CantCampos + 1
             Cadena = Mid(Cadena, No_Hasta, Len(Cadena))
             TotalReg = Len(Cadena)
             No_Desde = 1: No_Hasta = 1
          Loop
          If Total_Alumnos = 0 Then
             ReDim CamposFile(CantCampos) As Campos_Tabla
             For I = 0 To CantCampos
                 CamposFile(I).Campo = "C" & Format(I, "00")
             Next I
          End If
          End If
        ' Comenzamos la subida de los Abonos
          Cadena = Cod_Field
          I = 0: No_Desde = 1: No_Hasta = 1
          Cadena = Cod_Field
          TotalReg = Len(Cadena)
          Do While Len(Cadena) > 0
             Do
               No_Hasta = No_Hasta + 1
             Loop Until Mid(Cadena, No_Hasta, 1) = vbTab Or No_Hasta > TotalReg
             CamposFile(I).Valor = Mid(Cadena, No_Desde, No_Hasta)
             If Len(CamposFile(I).Valor) > 1 Then
               'Si el limitador de campos es un TAB
                If EsTab Then
                   If Mid(CamposFile(I).Valor, 1, 1) = vbTab Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 2, Len(CamposFile(I).Valor))
                   If Mid(CamposFile(I).Valor, Len(CamposFile(I).Valor), 1) = vbTab Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 1, Len(CamposFile(I).Valor) - 1)
                End If
             End If
             I = I + 1
             Cadena = Mid(Cadena, No_Hasta, Len(Cadena))
             TotalReg = Len(Cadena)
             No_Desde = 1: No_Hasta = 1
          Loop
         'Actualizamos de que alumnos vamos a ingresar el abono
          CodigoCli = Ninguno
          CodigoP = CStr(Val(CamposFile(7).Valor))
          FechaTexto = CamposFile(12).Valor
          With AdoAux.Recordset
              .MoveFirst
              .Find ("CI_RUC = '" & CodigoP & "' ")
               If Not .EOF Then CodigoCli = .Fields("Codigo")
          End With
          sSQL = "UPDATE Facturas " _
               & "SET X = 'F' " _
               & "WHERE CodigoC = '" & CodigoCli & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND TC NOT IN ('C','P') " _
               & "AND T <> 'A' "
          ConectarAdoExecute sSQL
         'Eliminamos los abonos de este dia
          MiFecha = BuscarFecha(FechaTexto)
          sSQL = "DELETE * " _
               & "FROM Trans_Abonos " _
               & "WHERE CodigoC = '" & CodigoCli & "' " _
               & "AND Fecha = #" & MiFecha & "# "
          ConectarAdoExecute sSQL
          
          Total_Alumnos = Total_Alumnos + 1
    Loop
  Close #NumFile
 'Actualizamos los saldos de las facturas
  Procesar_Saldo_De_Facturas FBancoPichincha, AdoAux, True
  ProgBarra.Max = Total_Alumnos + 1
 'Consultamos los alumnos a verificar
  sSQL = "SELECT F.*,C.Cliente,C.CI_RUC,C.Grupo " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.Saldo_MN > 0 " _
       & "AND F.TC NOT IN ('C','P') " _
       & "AND F.X = 'F' " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.Factura "
  SelectAdodc AdoAux, sSQL
  'MsgBox sSQL
' Comenzamos a leer el archivo de Abonos
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       RutaGeneraFile = UCase(Dir1.Path & "\" & NombreArchivo)
       Contador = 0: FileResp = 0
       NumFile = FreeFile
       Open RutaGeneraFile For Input As #NumFile
         Do While Not EOF(NumFile)
            Line Input #NumFile, Cod_Field
            ProgBarra.Value = ProgBarra.Value + 1
            No_Desde = 1: No_Hasta = 1
            Cadena = Cod_Field
            I = 0
            TotalReg = Len(Cadena)
            Do While Len(Cadena) > 0
               Do
                 No_Hasta = No_Hasta + 1
               Loop Until Mid(Cadena, No_Hasta, 1) = "," Or Mid(Cadena, No_Hasta, 1) = vbTab Or No_Hasta > TotalReg
               CamposFile(I).Valor = Mid(Cadena, No_Desde, No_Hasta)
               If Len(CamposFile(I).Valor) > 1 Then
                  If EsComa Then
                     If Mid(CamposFile(I).Valor, 1, 1) = "," Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 2, Len(CamposFile(I).Valor))
                     If Mid(CamposFile(I).Valor, Len(CamposFile(I).Valor), 1) = "," Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 1, Len(CamposFile(I).Valor) - 1)
                  End If
                  If EsTab Then
                     If Mid(CamposFile(I).Valor, 1, 1) = vbTab Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 2, Len(CamposFile(I).Valor))
                     If Mid(CamposFile(I).Valor, Len(CamposFile(I).Valor), 1) = vbTab Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 1, Len(CamposFile(I).Valor) - 1)
                  End If
               End If
               I = I + 1
               Cadena = Mid(Cadena, No_Hasta, Len(Cadena))
               TotalReg = Len(Cadena)
               No_Desde = 1: No_Hasta = 1
            Loop
          ' Obtenemos el Valor de Abonos de Pensiones o Matriculas
            CodigoCli = Ninguno
            CodigoP = CStr(Val(CamposFile(7).Valor))
            Total = Val(CamposFile(9).Valor) / 100
            MiFecha = CamposFile(12).Valor
            TextoCheque = Ninguno
            NombreBanco = "DEPOSITO EFECTIVO"
            If CamposFile(16).Valor <> "EFE" Then
               NombreBanco = "TRANS. " & Trim(CamposFile(16).Valor & ". " & CamposFile(17).Valor)
            End If
            With AdoAux.Recordset
                .MoveFirst
                .Find ("CI_RUC = '" & CodigoP & "' ")
                 If Not .EOF Then CodigoCli = .Fields("CodigoC")
            End With
            TextoCheque = Ninguno
            TotalIngreso = TotalIngreso + Total
           .MoveFirst
           .Find ("CI_RUC = '" & CodigoP & "' ")
            If Not .EOF Then
               CodigoCli = .Fields("CodigoC")
               TextoCheque = .Fields("Grupo")
            End If
            sSQL = "SELECT * " _
                 & "FROM Facturas " _
                 & "WHERE CodigoC = '" & CodigoCli & "' " _
                 & "AND Saldo_MN > 0 " _
                 & "AND TC NOT IN ('C','P') " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND T <> 'A' " _
                 & "ORDER BY Factura "
            SelectAdodc AdoAct, sSQL
            'MsgBox CodigoP & " - " & Total
            If AdoAct.Recordset.RecordCount > 0 Then
               Abono = 1
               Do While Not AdoAct.Recordset.EOF And Abono > 0
                  Abono = 0
                  TipoFactura = AdoAct.Recordset.Fields("TC")
                  Saldo = AdoAct.Recordset.Fields("Saldo_MN")
                  SaldoTotal = AdoAct.Recordset.Fields("Saldo_MN")
                  Factura_No = AdoAct.Recordset.Fields("Factura")
                  If Total <= Saldo Then
                     If Total > 0 Then Abono = Total
                  Else
                     Abono = Saldo
                  End If
                 'MsgBox CodigoCli & " - " & Total_ME & " - Fact = " & Factura_No & " - Abono = " & Abono
                  If Abono > 0 Then
                     SetAdoAddNew "Trans_Abonos"
                     SetAdoFields "T", Cancelado
                     SetAdoFields "TP", TipoFactura
                     SetAdoFields "CodigoC", CodigoCli
                     SetAdoFields "Fecha", MiFecha
                     SetAdoFields "Comprobante", Ninguno
                     SetAdoFields "Factura", Factura_No
                     SetAdoFields "Abono", Abono
                     SetAdoFields "Banco", NombreBanco
                     SetAdoFields "Cheque", TextoCheque
                     SetAdoFields "Cta", Cta_CajaG
                     SetAdoFields "Cta_CxP", Cta_Cobrar
                     SetAdoUpdate
                     SaldoTotal = SaldoTotal - Abono
                     If SaldoTotal <= 0 Then
                        sSQL = "UPDATE Facturas " _
                             & "SET Saldo_MN = " & SaldoTotal & ", T = 'C'  " _
                             & "WHERE CodigoC = '" & CodigoCli & "' " _
                             & "AND Factura = " & Factura_No & " " _
                             & "AND TC = '" & TipoFactura & "' " _
                             & "AND Periodo = '" & Periodo_Contable & "' " _
                             & "AND Item = '" & NumEmpresa & "' "
                        ConectarAdoExecute sSQL
                     End If
                End If
                Total = Total - Abono
                AdoAct.Recordset.MoveNext
               Loop
            Else
               TextoImprimio = TextoImprimio & CodigoP & vbCrLf
            End If
            Contador = Contador + 1
            If Contador > ProgBarra.Max Then Contador = ProgBarra.Max
            ProgBarra.Value = Contador
         Loop
         Procesar_Saldo_De_Facturas FBancoPichincha, AdoAux, True
         ProgBarra.Value = ProgBarra.Max
       Close #NumFile
   Else
       MsgBox "No Existen Facturas Pendientes"
   End If
  End With
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  MsgBox "ARCHIVO DE ABONO DEL DIA: " & FechaTexto & vbCrLf & vbCrLf _
       & "SE ACTUALIZARON: " & Total_Alumnos & " ALUMNOS." & vbCrLf & vbCrLf _
       & "EL CIERRE DIARIO DE CAJA ES POR " & Moneda & " " & Format(TotalIngreso, "#,##0.00") & vbCrLf & vbCrLf _
       & "OBTENIDO DE ARCHIVO: " & vbCrLf & vbCrLf & RutaGeneraFile
  If TextoImprimio <> "" Then FInfoError.Show
End Sub

Private Sub Command2_Click()
  Unload FBancoPichincha
End Sub

Private Sub Command7_Click()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim CamposFile() As Campos_Tabla
Dim EsComa As Boolean
Dim EsTab As Boolean
  FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  EsComa = False
  EsTab = False
  FechaValida MBFechaI
  MiFecha = BuscarFecha(MBFechaI)
  MiMes = Format(Month(MBFechaI), "00")
  FechaFin = BuscarFecha(UltimoDiaMes(MBFechaI))
  
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
  
  sSQL = "SELECT F.CodigoC,C.Actividad,C.Cliente,CI_RUC,Direccion,SUM(Saldo_MN) As Saldo_Pend " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND NOT F.TC IN ('C','P') " _
       & "AND F.T = 'P' " _
       & "AND LEN(C.Actividad) <= 1 " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.CodigoC = C.Codigo " _
       & "GROUP BY F.CodigoC,C.Actividad,C.Cliente,CI_RUC,Direccion " _
       & "HAVING SUM(Saldo_MN) > 0 " _
       & "ORDER BY F.CodigoC,C.Actividad,C.Cliente,CI_RUC,Direccion "
  SelectAdodc AdoAct, sSQL
' Comenzamos a generar el archivo: SCRECXX.TXT
  EsComa = True
  EsTab = False
  Contador = 0
  Factura_No = 0
  NumFile = FreeFile
  RutaGeneraFile = UCase(Dir1.Path & "\NOMINA\SCREC" & Month(MBFechaI) & ".TXT")
  MsgBox RutaGeneraFile
  Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
  With AdoAct.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .Fields("CodigoC")
          NombreCliente = .Fields("Cliente")
          Factura_No = Factura_No + 1
          Total = .Fields("Saldo_Pend")
          CodigoP = .Fields("CI_RUC")
          DireccionCli = .Fields("Direccion")
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = Trim(Mid(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = Trim(SinEspaciosDer(DireccionCli))
          Codigo1 = Format(Mid(GrupoNo, 1, 1), "00")
          
          Print #NumFile, "CO" & vbTab;
          Print #NumFile, "8540031" & vbTab;
          Print #NumFile, Contador & vbTab;
          Print #NumFile, Format(Factura_No, "0000000000") & vbTab;
          Print #NumFile, CodigoP & vbTab;
          Print #NumFile, "USD" & vbTab;
          Saldo = Fix(Total)
          NumStrg = Format(Saldo, "00000000000") & Format(Total - Saldo, "00")
          Print #NumFile, NumStrg & vbTab;
          Print #NumFile, "REC" & vbTab;
          Print #NumFile, "10" & vbTab;
          Print #NumFile, vbTab;     'CTE/AHO
          Print #NumFile, "0" & vbTab;     'No. Cta Cte/Aho
          Print #NumFile, "R" & vbTab;
          Print #NumFile, RUC & vbTab;
          Print #NumFile, Mid(Month(MBFechaI) & " " & NombreCliente, 1, 40) & vbTab;
          Print #NumFile, vbTab;
          Print #NumFile, vbTab;
          Print #NumFile, vbTab;
          Print #NumFile, vbTab;
          Print #NumFile, Month(MBFechaI) & vbTab;
          Print #NumFile, "Pensión Acumulada" & vbTab;
          Saldo = Fix(Total)
          NumStrg = Format(Saldo, "00000000000") & Format(Total - Saldo, "00")
          Print #NumFile, NumStrg & vbTab
         .MoveNext
       Loop
   End If
  End With
  Close #NumFile
' Comenzamos a generar el archivo: SCCOB.TXT
  EsComa = True
  EsTab = False
  Mes = Month(MiFecha)
  Anio = Val(Mid(Format(Year(MiFecha), "0000"), 2, 3))
  Dia = "15"
  sSQL = "SELECT F.CodigoC,C.Actividad,C.Cliente,CI_RUC,Direccion,SUM(Saldo_MN) As Saldo_Pend " _
       & "FROM Facturas As F,Clientes As C " _
       & "WHERE F.Item = '" & NumEmpresa & "' " _
       & "AND NOT F.TC IN ('C','P') " _
       & "AND F.T = 'P' " _
       & "AND LEN(C.Actividad) > 3 " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.CodigoC = C.Codigo " _
       & "GROUP BY F.CodigoC,C.Actividad,C.Cliente,CI_RUC,Direccion " _
       & "HAVING SUM(Saldo_MN) > 0 " _
       & "ORDER BY F.CodigoC,C.Actividad,C.Cliente,CI_RUC,Direccion "
  SelectAdodc AdoAct, sSQL
  
  NumFile = FreeFile
  RutaGeneraFile = UCase(Dir1.Path & "\NOMINA\SCCOB" & Month(MBFechaI) & ".TXT")
  MsgBox RutaGeneraFile
  Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
  With AdoAct.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Contador = Contador + 1
          CodigoCli = .Fields("CodigoC")
          NombreCliente = .Fields("Cliente")
          Factura_No = Factura_No + 1
          Total = .Fields("Saldo_Pend")
          CodigoP = .Fields("CI_RUC")
          DireccionCli = .Fields("Direccion")
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = Trim(Mid(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = Trim(SinEspaciosDer(DireccionCli))
          Codigo1 = Format(Mid(GrupoNo, 1, 1), "00")
          
          Print #NumFile, "CO" & vbTab;
          Print #NumFile, "8540031" & vbTab;
          Print #NumFile, Contador & vbTab;
          Print #NumFile, Format(Factura_No, "0000000000") & vbTab;
          Print #NumFile, CodigoP & vbTab;
          Print #NumFile, "USD" & vbTab;
          Saldo = Fix(Total)
          NumStrg = Format(Saldo, "00000000000") & Format(Total - Saldo, "00")
          Print #NumFile, NumStrg & vbTab;
          Print #NumFile, "CTA" & vbTab;
          Print #NumFile, "10" & vbTab;
          NumStrg = SinEspaciosIzq(.Fields("Actividad"))
          If Len(NumStrg) = 3 Then
             Print #NumFile, SinEspaciosIzq(.Fields("Actividad")) & vbTab;      'CTE/AHO
             Print #NumFile, SinEspaciosDer(.Fields("Actividad")) & vbTab;     'No. Cta Cte/Aho
          Else
             Print #NumFile, vbTab;       'CTE/AHO
             Print #NumFile, vbTab;      'No. Cta Cte/Aho
          End If
          Print #NumFile, "R" & vbTab;
          Print #NumFile, RUC & vbTab;
          Print #NumFile, Mid(Month(MBFechaI) & " " & NombreCliente, 1, 40) & vbTab;
          Print #NumFile, vbTab;
          Print #NumFile, vbTab;
          Print #NumFile, vbTab;
          Print #NumFile, vbTab;
          Print #NumFile, Month(MBFechaI) & vbTab;
          Print #NumFile, "Pensión Acumulada" & vbTab;
          Saldo = Fix(Total)
          NumStrg = Format(Saldo, "00000000000") & Format(Total - Saldo, "00")
          Print #NumFile, NumStrg & vbTab
         .MoveNext
       Loop
   End If
  End With
  Close #NumFile
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  FBancoPichincha.Caption = "FACTURACION DE BANCO INTERNACIONAL"
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
  SiguienteControl
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
  NombreArchivo = File1.FileName
  If KeyCode = vbKeyDelete Then
     Mensajes = "Esta seguro de Eliminar: " & File1.FileName
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then Kill File1.Path & "\" & File1.FileName
     File1.FileName = Dir1.Path & "\*.*"
  End If
End Sub

Private Sub File1_LostFocus()
  NombreArchivo = UCase(File1.FileName)
End Sub

Private Sub Form_Activate()
  FechaValida MBFechaI
  Drive1.Drive = Mid(RutaSysBases, 1, 2)
  RatonNormal
  RutaBackup = RutaSysBases & "\BANCO"
  TipoProcesos ""
  TipoFactura = "FA"
  FBancoPichincha.Caption = "FACTURACION DE BANCOS"
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> 0 " _
       & "ORDER BY CI_RUC "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     AdoAux.Recordset.MoveLast
     Codigo = AdoAux.Recordset.Fields("CI_RUC")
  End If
  sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd,* " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY Codigo_Inv "
  SelectDBCombo DCInv, AdoInv, sSQL, "NomProd"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FBancoPichincha
  ConectarAdodc AdoAux
  ConectarAdodc AdoAct
  ConectarAdodc AdoInv
  ConectarAdodc AdoQuery
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

