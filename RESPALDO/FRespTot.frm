VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form RespaldoTotal 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RESPALDO/RESTAURACION TOTAL"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir"
      Height          =   645
      Left            =   8085
      TabIndex        =   6
      Top             =   2625
      Width           =   1590
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Restaurar"
      Height          =   645
      Left            =   8085
      TabIndex        =   7
      Top             =   1890
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Descomprimir"
      Height          =   645
      Left            =   8085
      TabIndex        =   8
      Top             =   1155
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Respaldar"
      Height          =   645
      Left            =   8085
      TabIndex        =   9
      Top             =   420
      Width           =   1590
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   5565
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
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
   Begin VB.CheckBox CheqItem 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Solo Empresa Actual"
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
      Left            =   6930
      TabIndex        =   3
      Top             =   105
      Width           =   2220
   End
   Begin VB.CheckBox CheqFechas 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Fechas de Restauración o Respaldos Total:"
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
      TabIndex        =   0
      Top             =   105
      Width           =   4110
   End
   Begin VB.ListBox LstTablas 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   4530
   End
   Begin MSAdodcLib.Adodc AdoRespaldo 
      Height          =   330
      Left            =   315
      Top             =   2835
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Adodc1"
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
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFC0C0&
      Height          =   3795
      Left            =   4725
      TabIndex        =   5
      Top             =   525
      Width           =   3270
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   1890
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
   Begin MSAdodcLib.Adodc AdoAct 
      Height          =   330
      Left            =   735
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
      Left            =   4305
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
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
   Begin MSComDlg.CommonDialog CDialogDir 
      Left            =   7035
      Top             =   2415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "RespaldoTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AdoStrCnnOld As String
Dim AdoStrCnn1 As String
Dim PathEmpresa1 As String
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim XAdoStrCnn As String
Dim IJ As Long
Dim ModuloResp As String
Dim RetVal
Dim Cont1  As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
Dim NombreFileZip As String

Public Sub LeerCamposTablaT(NumFile As Integer)
Dim CodBusp1 As Variant
    Cod_Emp = "": Cod_Base = "": Cod_Field = ""
    Line Input #NumFile, Cod_Field
   'MsgBox Cod_Field
    No_Desde = 1: No_Hasta = 1
    Cadena = Cod_Field
    For I = 1 To CantCamposT
        Do
           No_Hasta = No_Hasta + 1
        Loop Until Mid(Cadena, No_Hasta, 1) = vbTab Or No_Hasta >= Len(Cod_Field)
        If No_Hasta >= Len(Cod_Field) Then No_Hasta = 1
        TipoC(I).Valor = Trim(Mid(Cadena, No_Desde, No_Hasta - 1))
        Cadena = Mid(Cadena, No_Hasta + 1, Len(Cadena))
        No_Desde = 1: No_Hasta = 1
    Next I
End Sub

Public Sub AbrirCamposSQLT(NumFile As Integer)
    Cod_Emp = "": Cod_Base = "": Cod_Field = ""
    'Line Input #NumFile, Cod_Base
    'TotalReg = 0
    'Cod_Base = SinEspaciosDer(Cod_Base)
    Line Input #NumFile, Cod_Field
    'MsgBox Cod_Base & vbCrLf & Cod_Field
    CantCamposT = 0
    For I = 1 To Len(Cod_Field)
        If Mid(Cod_Field, I, 1) = vbTab Then CantCamposT = CantCamposT + 1
    Next I
    ReDim TipoC(CantCamposT) As Campos_Tabla
    No_Desde = 1: No_Hasta = 1
    Cadena = Cod_Field
    For I = 1 To CantCamposT
        Do
           No_Hasta = No_Hasta + 1
        Loop Until Mid(Cadena, No_Hasta, 1) = vbTab
        TipoC(I).Campo = Trim(Mid(Cadena, No_Desde, No_Hasta - 1))
        Cadena = Mid(Cadena, No_Hasta + 1, Len(Cadena))
        No_Desde = 1: No_Hasta = 1
    Next I
End Sub

Private Sub CheqFechas_Click()
   If CheqFechas.value = 1 Then
      MBFechaI.Visible = True
      MBFechaF.Visible = True
   Else
      MBFechaI.Visible = False
      MBFechaF.Visible = False
   End If
End Sub

Private Sub Command1_Click()
    Empaquetar_Archivos_Zip
    Unload RespaldoTotal
End Sub

Private Sub Command2_Click()
Dim IdFile As Byte

  RutaOrigen = UCase(SelectZipFile(CDialogDir, OpenZip))
  If RutaOrigen <> "" Then
     IdFile = InStrRev(RutaOrigen, "\")
     NombreFileZip = Mid(RutaOrigen, IdFile + 1, Len(RutaOrigen))
     IdFile = Len(RutaOrigen)
     Codigo4 = Mid(RutaOrigen, IdFile - 2, 3)
    'MsgBox RutaOrigen & vbCrLf & Codigo1 & vbCrLf & Codigo2 & vbCrLf & Codigo3 & vbCrLf & Codigo4 & vbCrLf & NombreFileZip
     If Codigo4 = "ZIP" Then
        RatonReloj
        ConSubDir = False
        Contador = 0: FileResp = 0
      ' Eliminamos archivos de otros dias
        File1.Filename = RutaSysBases & "\DATOS\TOTAL\*.BDD"
        File1.Refresh
        If File1.ListCount > 0 Then Kill RutaSysBases & "\DATOS\TOTAL\*.BDD"
        RutaDestino = RutaSysBases & "\DATOS\TOTAL\"
        MBFechaI = Format(Replace(Mid(NombreFileZip, Len(NombreFileZip) - 14, 11), "-", "/"), FormatoFechas)
        FechaValida MBFechaI
       'Pasamos a descomprimir
        UnZip RutaOrigen, RutaDestino
        File1.Refresh
        RatonNormal
        MsgBox "FIN DEL PROCESO DE DESCOMPRESION," & vbCrLf & vbCrLf _
             & "PROCEDA A RESTAURAR LA INFORMACION"
        RespaldoTotal.Caption = "MODULO DE RESPALDO TOTAL"
        'CommandButton2.SetFocus
     Else
        MsgBox "ESTE ARCHIVO NO ES VALIDO," & vbCrLf & vbCrLf _
             & "NO SE PUEDE PROCESAR"
     End If
  End If
End Sub

Private Sub Command3_Click()
Dim AuxNumEmp As String
Dim NombreTabla  As String
Dim NumReg As Long
Dim XItem As Boolean
Dim XInserta As Boolean
Dim ContReg As Long
  RatonReloj
  Progreso_Barra.Mensaje_Box = "PROGRESO DE LA RESTAURACION"
  Progreso_Iniciar
  MiTiempo = Time
  XItem = False
  AuxNumEmp = NumEmpresa
 'If CheqItem.value = 1 Then NumEmpresa = Mid(NombreFileZip, 1, 3)
  LstTablas.Clear
  File1.Filename = RutaSysBases & "\DATOS\TOTAL\*.BDD"
  File1.Refresh
  Progreso_Barra.Valor_Maximo = File1.ListCount + 10
  For Contador = 0 To File1.ListCount - 1
      TotalReg = 0
      ContReg = 0
      RutaGeneraFile = RutaSysBases & "\DATOS\TOTAL\" & File1.List(Contador)
      NumFile = FreeFile
      Open RutaGeneraFile For Input As #NumFile
           Line Input #NumFile, Cod_Field
           NombreTabla = SinEspaciosDer(Cod_Field)
           TotalReg = Val(SinEspaciosIzq(Trim(Mid$(Cod_Field, 11, Len(Cod_Field)))))
           XItem = False
           Progreso_Barra.Mensaje_Box = "Restaurando: " & NombreTabla
           Progreso_Esperar
           
           sSQL = "SELECT TOP 1 * " _
                & "FROM " & NombreTabla & " "
           SelectAdodc AdoRespaldo, sSQL
           For I = 0 To AdoRespaldo.Recordset.Fields.Count - 1
               If AdoRespaldo.Recordset.Fields(I).Name = "Item" And CheqItem.value = 1 Then XItem = True
           Next I
           
           Select Case NombreTabla
             Case "Accesos", "Clientes"
                 'No hacemos nada por ahora
             Case Else
                  sSQL = "DELETE * " _
                       & "FROM " & NombreTabla & " "
                  If XItem Then sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
                  ConectarAdoExecute sSQL
                 'MsgBox sSQL
           End Select
           
           If XItem And CheqItem.value <> 1 Then
              sSQL = "DELETE * " _
                   & "FROM " & NombreTabla & " " _
                   & "WHERE Item <> '000' "
              ConectarAdoExecute sSQL
           End If
          'Abre los campos que vamos a subir
           AbrirCamposSQLT NumFile
           Do While Not EOF(NumFile)
              Progreso_Barra.Mensaje_Box = "T: " & Format(Time - MiTiempo, "HH:MM:SS") _
                                         & " Restaurando: " & NombreTabla & " -> " _
                                         & Format(ContReg / TotalReg, "00%")
              Progreso_Esperar True
              XInserta = True
              LeerCamposTablaT NumFile
              Select Case NombreTabla
                Case "Accesos", "Clientes"
                     For I = 1 To CantCamposT
                         If TipoC(I).Campo = "Codigo" Then
                            sSQL = "SELECT * " _
                                 & "FROM " & NombreTabla & " " _
                                 & "WHERE Codigo = '" & TipoC(I).Valor & "' "
                            If XItem Then sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
                            SelectAdodc AdoAux, sSQL
                            If AdoAux.Recordset.RecordCount > 0 Then XInserta = False
                            'MsgBox XInserta & vbCrLf & vbCrLf & sSQL & vbCrLf & vbCrLf & AdoAux.Recordset.RecordCount
                         End If
                     Next I
              End Select
              
              If XInserta Then
                 SetAdoAddNew NombreTabla, True
                 Cadena = ""
                 For I = 1 To CantCamposT
                     TipoC(I).Valor = Replace(TipoC(I).Valor, Chr(17) & Chr(31), vbCrLf)
                     TipoC(I).Valor = Replace(TipoC(I).Valor, "'", "`")
                     SetAdoFields TipoC(I).Campo, TipoC(I).Valor
                     Cadena = Cadena & TipoC(I).Campo & " = " & TipoC(I).Valor & vbCrLf
                 Next I
                 SetAdoUpdate
                'MsgBox UCase$(NombreTabla) & vbCrLf & Cadena
              End If
              ContReg = ContReg + 1
           Loop
      Close #NumFile
      LstTablas.AddItem "Reg." & Space(8 - Len(CStr(TotalReg))) & TotalReg & " - " & NombreTabla
      LstTablas.Refresh
     'MsgBox "Terminado...."
  Next Contador
  Progreso_Barra.Mensaje_Box = "Reindexando Registros"
  Progreso_Esperar
  sSQL = "UPDATE Transacciones " _
       & "SET Procesado = 0 "
  If CheqItem.value <> 0 Then sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
  ConectarAdoExecute sSQL
  
  Progreso_Esperar
  sSQL = "UPDATE Trans_Kardex " _
       & "SET Procesado = 0 "
  If CheqItem.value <> 0 Then sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
  ConectarAdoExecute sSQL
  
  Progreso_Barra.Mensaje_Box = "Proceso Terminado"
  Progreso_Esperar
  Progreso_Final
  RatonNormal
  MsgBox "Proceso Terminado"
End Sub

Private Sub Command4_Click()
  Unload RespaldoTotal
End Sub

Private Sub Form_Activate()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim IJ As Long
Dim IdTime As Long
Dim strCnn As String
  RatonReloj
' Consultamos las cuentas de la tabla
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.Open AdoStrCnn
  Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
  Do Until RstSchema.EOF
     If RstSchema!TABLE_TYPE = "TABLE" And Mid(RstSchema!TABLE_NAME, 1, 1) <> "~" Then LstTablas.AddItem RstSchema!TABLE_NAME
     RstSchema.MoveNext
  Loop
  RstSchema.Close
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm RespaldoTotal
  ConectarAdodc AdoAux
  ConectarAdodc AdoAct
  ConectarAdodc AdoRespaldo
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

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Public Sub Respaldar_Bases()
Dim CaptionOld As String
Dim NombreFile As String
Dim NombreTabla  As String
Dim CadFileReg As String
Dim ContadorReg As Long
Dim TotalCampo As Integer
Dim ValorBool As String
Dim Numero_Tabla, I As Long
Dim TextoFileEmp As String
Dim XItem As Boolean
Dim XFecha As Boolean
Dim XPeriodo As Boolean
Dim XRespaldar As Boolean
Dim XWhere As Boolean
RatonReloj
MiTiempo = Time
TextoFileEmp = ""
Numero_Tabla = 0
    
  sSQL = "UPDATE Fechas_Balance " _
       & "SET Fecha_Inicial = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha_Inicial <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Fechas_Balance " _
       & "SET Fecha_Final = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha_Final <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Comprobantes " _
       & "SET Fecha = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Detalle_Factura " _
       & "SET Fecha = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Facturas " _
       & "SET Fecha = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Trans_Abonos " _
       & "SET Fecha = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Trans_SubCtas " _
       & "SET Fecha = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Trans_SubCtas " _
       & "SET Fecha_V = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha_V <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Transacciones " _
       & "SET Fecha = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
  sSQL = "UPDATE Transacciones " _
       & "SET Fecha_Efec = #" & BuscarFecha("01/01/2000") & "# " _
       & "WHERE Fecha_Efec <= #" & BuscarFecha("01/01/1900") & "# "
  ConectarAdoExecute sSQL
  
FechaIni = BuscarFecha(MBFechaI)
FechaFin = BuscarFecha(MBFechaF)

HoraSistema = Format(Time, "HH:MM:SS")
'Creamos un archivo para poder limpiar los datos antiguos
NumFile = FreeFile
RutaGeneraFile = RutaSysBases & "\DATOS\TOTAL\U" & CodigoUsuario & ".BDD"
Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
Close #NumFile
Kill RutaSysBases & "\DATOS\TOTAL\*.BDD"
For Contador = 0 To LstTablas.ListCount - 1
    NombreTabla = LstTablas.List(Contador)
    HoraSistema = Format(Time, "HH:MM:SS")
    RespaldoTotal.Caption = HoraSistema & " =>> Respaldando: " & NombreTabla
    XRespaldar = True
    If Mid(NombreTabla, 1, 4) = "Tipo" Then XRespaldar = False
    If Mid(NombreTabla, 1, 5) = "Tabla" Then XRespaldar = False
    If Mid(NombreTabla, 1, 7) = "Asiento" Then XRespaldar = False
    If XRespaldar Then
       LstTablas.Text = NombreTabla
       RespaldoTotal.Caption = " Respaldo Programado - Tiempo: " & Format(Time - MiTiempo, "HH:MM:SS") & " - " & NombreTabla
      'Abrimos la tabla a respaldar
       XItem = False
       XFecha = False
       XWhere = False
       XPeriodo = False
       sSQL = "SELECT TOP 1 * " _
            & "FROM " & NombreTabla & " "
       SelectAdodc AdoRespaldo, sSQL
       For I = 0 To AdoRespaldo.Recordset.Fields.Count - 1
           If AdoRespaldo.Recordset.Fields(I).Name = "Item" Then XItem = True
           If AdoRespaldo.Recordset.Fields(I).Name = "Fecha" Then XFecha = True
           If AdoRespaldo.Recordset.Fields(I).Name = "Periodo" Then XPeriodo = True
       Next I
       sSQL = "SELECT * " _
            & "FROM " & NombreTabla & " "
       If XItem And CheqItem.value = 1 Then
          sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
          XWhere = True
       End If
        
       If XFecha And CheqFechas.value = 1 Then
          If XWhere Then
             sSQL = sSQL & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
          Else
             sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
             XWhere = True
          End If
       End If
       If XItem And XPeriodo And XFecha Then
          sSQL = sSQL & "ORDER BY Item,Periodo,Fecha "
       ElseIf XItem And XPeriodo Then
          sSQL = sSQL & "ORDER BY Item,Periodo "
       ElseIf XItem Then
          sSQL = sSQL & "ORDER BY Item "
       ElseIf XFecha Then
          sSQL = sSQL & "ORDER BY Fecha "
       End If
      'MsgBox "SQL > " & vbCrLf & sSQL
       SelectAdodc AdoRespaldo, sSQL
       ContadorReg = 0
       If FileResp <= 0 Then FileResp = 1
       With AdoRespaldo.Recordset
        If .RecordCount > 0 Then
           'Abrimos el archivo de respaldo
            Numero_Tabla = Numero_Tabla + 1
            NombreFile = "Dato_" & Format(Numero_Tabla, "000") & ".BDD"
            TextoFileEmp = TextoFileEmp & vbCrLf & NombreFile & " => " & NombreTabla
            RutaGeneraFile = RutaSysBases & "\DATOS\TOTAL\" & NombreFile
            NumFile = FreeFile
            Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
            
           'Averiguamos en numero de registros
           .MoveLast
            TotalReg = .RecordCount
            TotalCampo = .Fields.Count - 1
            Print #NumFile, HoraSistema & " - " & Format(TotalReg, "##0") & " - " & NombreTabla
            ReDim TipoC(TotalCampo) As Campos_Tabla
            CadFileReg = ""
            For I = 0 To TotalCampo
                TipoC(I).Campo = .Fields(I).Name
                TipoC(I).Ancho = AnchoTipoCampoTexto(.Fields(I))
                CadFileReg = CadFileReg & TipoC(I).Campo & vbTab      '"|"
            Next I
            Print #NumFile, CadFileReg
            ContadorReg = 0
           .MoveFirst
            Do While Not .EOF
               CadFileReg = ""
               ContadorReg = ContadorReg + 1
               RespaldoTotal.Caption = " Dia Programado - Tiempo: " & Format(Time - MiTiempo, "HH:MM:SS") & " - " & Format(Contador / LstTablas.ListCount, "00%") & " - " & NombreTabla _
                              & " - Datos (" & Format(ContadorReg / .RecordCount, "00%") & ")"
               For I = 0 To TotalCampo
                   TipoC(I).Valor = Ninguno
                   If IsNull(.Fields(I)) Or IsEmpty(.Fields(I)) Then
                      TipoC(I).Valor = "0"
                   Else
                      Select Case .Fields(I).Type
                        Case TadDate, TadDate1
                             TipoC(I).Valor = .Fields(I)
                        Case TadBoolean
                             TipoC(I).Valor = CInt(.Fields(I))
                        Case TadByte, TadInteger, TadLong, TadSingle, TadDouble, TadCurrency
                             TipoC(I).Valor = .Fields(I)
                        Case TadText
                             TipoC(I).Valor = .Fields(I)
                             If UCase(.Fields(I).Name) = "RUC_CI" Then TipoC(I).Valor = CompilarRUC_CI(.Fields(I))
                             If UCase(.Fields(I).Name) = "RUC" Then TipoC(I).Valor = CompilarRUC_CI(.Fields(I))
                             If UCase(.Fields(I).Name) = "CI" Then TipoC(I).Valor = CompilarRUC_CI(.Fields(I))
                             If UCase(.Fields(I).Name) = "CEDULA" Then TipoC(I).Valor = CompilarRUC_CI(.Fields(I))
                             If UCase(.Fields(I).Name) = "ITEM" Then TipoC(I).Valor = Format(.Fields(I), "000")
                        Case Else
                             TipoC(I).Valor = .Fields(I)
                      End Select
                     'Reemplazamos los Return por otros codigo ASCII
                      TipoC(I).Valor = Replace(TipoC(I).Valor, vbCrLf, Chr(17) & Chr(31))
                   End If
              Next I
              For I = 0 To TotalCampo
                  CadFileReg = CadFileReg & TipoC(I).Valor & vbTab    '"|"
              Next I
              Print #NumFile, CadFileReg
             .MoveNext
            Loop
           .MoveFirst
            FileResp = FileResp + 1
            Close #NumFile
        End If
       End With
    End If
Next Contador
RatonNormal
End Sub

Public Sub Empaquetar_Archivos_Zip()
Dim RutaZip As String
Dim FileZip As String
Dim BorrarFileZip As Boolean
Dim Resultado As Long
Dim intContadorFicheros As Integer
Dim FuncionesZip As ZIPUSERFUNCTIONS
Dim OpcionesZip As ZPOPT

    FechaValida MBFechaI
    FechaValida MBFechaF
    FechaIni = BuscarFecha(MBFechaI)
    FechaFin = BuscarFecha(MBFechaF)
    
    NoAnio = Year(FechaSistema)
    NoMes = Month(FechaSistema)
    NoDias = Day(FechaSistema)
    
    FuncionesZip.DLLComment = DevolverDireccionMemoria(AddressOf FuncionParaProcesarComentarios)
    FuncionesZip.DLLPassword = DevolverDireccionMemoria(AddressOf FuncionParaProcesarPassword)
    FuncionesZip.DLLPrnt = DevolverDireccionMemoria(AddressOf FuncionParaProcesarMensajes)
    FuncionesZip.DLLService = DevolverDireccionMemoria(AddressOf FuncionParaProcesarServicios)
    
    FileZip = NumEmpresa & " Respaldo Total " & UCase(Replace(Format(FechaSistema, "yyyy/MM/dd"), "/", "-")) & ".ZIP"
    RutaZip = RutaSysBases & "\DATOS\TOTAL\"
    RutaOrigen = RutaZip & FileZip
    RutaDestino = RutaZip & "*.BDD"
    BorrarFileZip = False
    Cadena = Dir(RutaZip, vbNormal)  'Recupera la primera entrada.
    Do While Cadena <> ""
       If Cadena <> "." And Cadena <> ".." Then
          If (GetAttr(RutaZip & Cadena) And vbNormal) = vbNormal Then
             If Cadena = FileZip Then BorrarFileZip = True
          End If
       End If
       Cadena = Dir
    Loop
    
    Cadena = Dir(RutaOrigen, vbArchive)
    If BorrarFileZip Then Kill RutaZip & FileZip
  
    Respaldar_Bases
    ChDir RutaZip
    intContadorFicheros = 0
    Cadena = Dir(RutaZip, vbNormal)  'Recupera la primera entrada.
    Do While Cadena <> ""
       If Cadena <> "." And Cadena <> ".." Then
          If (GetAttr(RutaZip & Cadena) And vbNormal) = vbNormal Then
             If Mid(Cadena, Len(Cadena) - 2.3) = "BDD" Then
                NombresFicherosZip.S(intContadorFicheros) = Cadena
                intContadorFicheros = intContadorFicheros + 1
             End If
          End If
       End If
       Cadena = Dir
    Loop
    NombreArchivoZip = RutaOrigen
    Resultado = ZpInit(FuncionesZip)
    Resultado = ZpSetOptions(OpcionesZip)
    Resultado = ZpArchive(intContadorFicheros, NombreArchivoZip, NombresFicherosZip)
    ChDir RutaEmpresa
    RatonNormal
    MsgBox "Respaldo Completo de la base de Datos listo y empaquetado"
End Sub

