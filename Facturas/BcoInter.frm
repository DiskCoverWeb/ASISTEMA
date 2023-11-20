VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FBancoInternacional1 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BANCO INTERNACIONAL"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "BcoInter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "BcoInter.frx":0442
   ScaleHeight     =   4530
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "BcoInter.frx":168F
      DataSource      =   "AdoInv"
      Height          =   315
      Left            =   2415
      TabIndex        =   2
      Top             =   945
      Width           =   4215
      _ExtentX        =   7435
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
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FFFF&
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
      Picture         =   "BcoInter.frx":16A4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   105
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
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
      Picture         =   "BcoInter.frx":1EAE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1155
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFFF&
      Height          =   2430
      Left            =   3360
      TabIndex        =   9
      Top             =   1575
      Width           =   3270
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   4095
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
      Top             =   1575
      Width           =   3165
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
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
      Picture         =   "BcoInter.frx":2754
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
      Height          =   2115
      Left            =   105
      TabIndex        =   6
      Top             =   1890
      Width           =   3165
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   525
      Top             =   2415
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
      Top             =   2730
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
      Top             =   3045
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
      Top             =   945
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   65535
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
      Top             =   3360
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
      BackColor       =   &H0000FFFF&
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
      Top             =   1365
      Width           =   3270
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
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
      Top             =   945
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
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
      Top             =   1365
      Width           =   3165
   End
End
Attribute VB_Name = "FBancoInternacional1"
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
          Cadena = Cod_Field
          FechaTexto = ""
          TotalReg = Len(Cadena)
          No_Desde = 1: No_Hasta = 1
          CantCampos = 0
          Do While Len(Cadena) > 0
             Do
               No_Hasta = No_Hasta + 1
               If Mid(Cadena, No_Hasta, 1) = "," Then EsComa = True
               If Mid(Cadena, No_Hasta, 1) = vbTab Then EsTab = True
             Loop Until Mid(Cadena, No_Hasta, 1) = "," Or Mid(Cadena, No_Hasta, 1) = vbTab Or No_Hasta > TotalReg
            'Obtenemos la fecha de subida
             If CantCampos = 8 Then FechaTexto = Trim(Mid(Cadena, No_Desde + 1, No_Hasta - 2)) & "/"
             If CantCampos = 9 Then FechaTexto = FechaTexto & Trim(Mid(Cadena, No_Desde + 1, No_Hasta - 2)) & "/20"
             If CantCampos = 10 Then FechaTexto = FechaTexto & Trim(Mid(Cadena, No_Desde + 1, No_Hasta - 2))
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
          Cadena = Cod_Field
          I = 0: No_Desde = 1: No_Hasta = 1
          Cadena = Cod_Field
          TotalReg = Len(Cadena)
          Do While Len(Cadena) > 0
             Do
               No_Hasta = No_Hasta + 1
             Loop Until Mid(Cadena, No_Hasta, 1) = "," Or Mid(Cadena, No_Hasta, 1) = vbTab Or No_Hasta > TotalReg
             CamposFile(I).Valor = Mid(Cadena, No_Desde, No_Hasta)
             If Len(CamposFile(I).Valor) > 1 Then
               'Si el limitador de campos es la coma
                If EsComa Then
                   If Mid(CamposFile(I).Valor, 1, 1) = "," Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 2, Len(CamposFile(I).Valor))
                   If Mid(CamposFile(I).Valor, Len(CamposFile(I).Valor), 1) = "," Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 1, Len(CamposFile(I).Valor) - 1)
                End If
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
          CodigoP = CamposFile(1).Valor
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
          Total_Alumnos = Total_Alumnos + 1
    Loop
  Close #NumFile
 'Eliminamos los abonos de este dia
  MiFecha = BuscarFecha(FechaTexto)
  sSQL = "DELETE * " _
       & "FROM Trans_Abonos " _
       & "WHERE Fecha = #" & MiFecha & "# "
  ConectarAdoExecute sSQL
 'Actualizamos los saldos de las facturas
  Procesar_Saldo_De_Facturas FBancoInternacional, AdoAux, True
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
            CodigoInv = CamposFile(11).Valor
            'CodigoP = Trim(Mid(CamposFile(1).Valor, 3, 8))
            CodigoP = Trim(CamposFile(1).Valor)
            CodigoCli = Ninguno
            TextoCheque = Ninguno
            Total = Val(CamposFile(32).Valor) + Val(CamposFile(55).Valor)
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
                     SetAdoFields "Banco", "DEPOSITO: (" & TextoCheque & ")"
                     SetAdoFields "Cheque", Format(Total_ME, "##0.00")
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
         Procesar_Saldo_De_Facturas FBancoInternacional, AdoAux, True
         ProgBarra.Value = ProgBarra.Max
       Close #NumFile
   Else
       MsgBox "No Existen Facturas Pendientes"
   End If
  End With
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  FBancoInternacional.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  MsgBox "ARCHIVO DE ABONO DEL DIA: " & FechaTexto & vbCrLf & vbCrLf _
       & "SE ACTUALIZARON: " & Total_Alumnos & " ALUMNOS." & vbCrLf & vbCrLf _
       & "EL CIERRE DIARIO DE CAJA ES POR " & Moneda & " " & Format(TotalIngreso, "#,##0.00") & vbCrLf & vbCrLf _
       & "OBTENIDO DE ARCHIVO: " & vbCrLf & vbCrLf & RutaGeneraFile
  If TextoImprimio <> "" Then FInfoError.Show
End Sub

Private Sub Command2_Click()
  Unload FBancoInternacional
End Sub

Private Sub Command7_Click()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim CamposFile() As Campos_Tabla
Dim EsComa As Boolean
Dim EsTab As Boolean
  FBancoInternacional.Caption = "FACTURACION DE BANCO INTERNACIONAL"
  EsComa = False
  EsTab = False
  FechaValida MBFechaI
  MiFecha = BuscarFecha(MBFechaI)
  TextoImprimio = ""
  AuxNumEmp = NumEmpresa
  Cta_Cobrar = Ninguno
  FechaTexto = FechaSistema
  
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " " _
       & "ORDER BY CI_RUC "
  SelectAdodc AdoAct, sSQL
  
  sSQL = "SELECT F.*,C.Grupo " _
       & "FROM Detalle_Factura As F,Clientes As C " _
       & "WHERE F.Fecha = #" & MiFecha & "# " _
       & "AND F.Item = '" & NumEmpresa & "' " _
       & "AND F.Periodo = '" & Periodo_Contable & "' " _
       & "AND F.T <> 'A' " _
       & "AND F.CodigoC = C.Codigo " _
       & "ORDER BY F.CodigoC,F.Codigo "
  SelectAdodc AdoAux, sSQL
' Comenzamos a generar el archivo: COALU.TXT
  EsComa = True
  EsTab = False
  NumFile = FreeFile
  RutaGeneraFile = UCase(Dir1.Path & "\COALU.TXT")
  MsgBox RutaGeneraFile
  Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
  With AdoAct.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoCli = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          CodigoP = Format(.Fields("CI_RUC"), "0000000000")
          DireccionCli = .Fields("Direccion")
          Codigo3 = SinEspaciosDer(DireccionCli)
          DireccionCli = Trim(Mid(DireccionCli, 1, Len(DireccionCli) - Len(Codigo3)))
          Codigo3 = Trim(SinEspaciosDer(DireccionCli))
          GrupoNo = .Fields("Grupo")
          Codigo1 = Format(Mid(GrupoNo, 1, 1), "00")
          Codigo2 = Format(Val(Mid(GrupoNo, 3, Len(GrupoNo) - 1)), "000") & Mid(GrupoNo, Len(GrupoNo), 1)
         'Empieza la trama
          Print #NumFile, Format(CodigoDelBanco, "0000");
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, CodigoP;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, NombreCliente;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, Codigo1;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, Codigo2;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, "DIURNO";
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, "A";
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, "0";
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, "0";
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, ",,";
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, Codigo3;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, "N";
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, "S" & Space(96)
         .MoveNext
       Loop
   End If
  End With
  Close #NumFile
' Comenzamos a generar el archivo: CODET.TXT
  EsComa = True
  EsTab = False
  Mes = Month(MiFecha)
  Anio = Val(Mid(Format(Year(MiFecha), "0000"), 2, 3))
  Dia = "15"
  NumFile = FreeFile
  RutaGeneraFile = UCase(Dir1.Path & "\CODET.TXT")
  MsgBox RutaGeneraFile
  Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
  With AdoAct.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          CodigoCli = .Fields("Codigo")
          NombreCliente = .Fields("Cliente")
          CodigoP = Format(.Fields("CI_RUC"), "0000000000")
          DireccionCli = .Fields("Direccion")
          GrupoNo = .Fields("Grupo")
          Codigo1 = Format(Mid(GrupoNo, 1, 1), "00")
          Codigo2 = Format(Val(Mid(GrupoNo, 3, Len(GrupoNo) - 1)), "000") & Mid(GrupoNo, Len(GrupoNo), 1)
         'Empieza la trama
          Print #NumFile, Format(CodigoDelBanco, "0000");
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, CodigoP;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, Mes;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, Anio;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Codigo3 = "E,15"
          Print #NumFile, Codigo3;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, Mes;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, Anio;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Codigo3 = "0,0,0"
          Print #NumFile, Codigo3;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Total = 0
          Contador = 0
          If AdoAux.Recordset.RecordCount > 0 Then
             AdoAux.Recordset.MoveFirst
             AdoAux.Recordset.Find ("CodigoC = '" & CodigoCli & "' ")
             If Not AdoAux.Recordset.EOF Then
                Si_No = True
                Do While Not AdoAux.Recordset.EOF And Si_No
                   If AdoAux.Recordset.Fields("CodigoC") = CodigoCli Then
                      Contador = Contador + 1
                      Codigos = AdoAux.Recordset.Fields("Codigo")
                      Codigos = Mid(Codigos, Len(Codigos) - 2, 3)
                      Print #NumFile, Codigos;
                      If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
                   Else
                      Si_No = False
                   End If
                   AdoAux.Recordset.MoveNext
                Loop
             End If
          End If
          For I = 0 To 10 - Contador - 1
           If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Next I
          If AdoAux.Recordset.RecordCount > 0 Then
             AdoAux.Recordset.MoveFirst
             AdoAux.Recordset.Find ("CodigoC = '" & CodigoCli & "' ")
             If Not AdoAux.Recordset.EOF Then
                Si_No = True
                Do While Not AdoAux.Recordset.EOF And Si_No
                   If AdoAux.Recordset.Fields("CodigoC") = CodigoCli Then
                      Total = Total + AdoAux.Recordset.Fields("Total")
                      Codigos = Format(AdoAux.Recordset.Fields("Total"), "0.00")
                      Print #NumFile, Codigos;
                      If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
                   Else
                      Si_No = False
                   End If
                   AdoAux.Recordset.MoveNext
                Loop
             End If
          End If
          For I = 0 To 10 - Contador - 1
           If EsComa Then Print #NumFile, "0.00,"; Else Print #NumFile, "0.00" & vbTab;
          Next I
          Print #NumFile, "USD";
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Codigos = Format(Total, "0.00")
          Print #NumFile, Codigos;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Codigo3 = "C,0.50,,,,,,0,0,0,,,0,,0,0,0,,0,,,,0.00,0.00,0"
          Print #NumFile, Codigo3;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, "0.00" & Space(82)
          'Print #NumFile, DireccionCli
         .MoveNext
       Loop
   End If
  End With
  Close #NumFile
  
  sSQL = "SELECT TC,Codigo_Inv,Producto " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY Codigo_Inv,Producto "
  SelectAdodc AdoAux, sSQL
' Comenzamos a generar el archivo: CORUB.TXT
  EsComa = True
  EsTab = False
  NumFile = FreeFile
  RutaGeneraFile = UCase(Dir1.Path & "\CORUB.TXT")
  MsgBox RutaGeneraFile
  Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
  With AdoAux.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          K = Len(.Fields("Codigo_Inv"))
          Codigo = Mid(.Fields("Codigo_Inv"), K - 2, 3)
          Producto = .Fields("Producto")
         'Empieza la trama
          Print #NumFile, Codigo;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, Producto;
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, Format(CodigoDelBanco, "0000");
          If EsComa Then Print #NumFile, ","; Else Print #NumFile, vbTab;
          Print #NumFile, "N,0.00,S"
         .MoveNext
       Loop
   End If
  End With
  Close #NumFile
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  FBancoInternacional.Caption = "FACTURACION DE BANCO INTERNACIONAL"
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
  FBancoInternacional.Caption = "FACTURACION DE BANCOS"
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
  CentrarForm FBancoInternacional
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

