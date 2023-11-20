VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIUpdate 
   BackColor       =   &H80000018&
   Caption         =   "VERSION DE JULIO-2022"
   ClientHeight    =   3900
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   8625
   Icon            =   "MDIUpdate.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIUpdate.frx":0ECA
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PictMDI 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   8625
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   8625
   End
   Begin VB.Timer Timer1 
      Left            =   525
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StaBarEmp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3525
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIUpdate.frx":27088
            Key             =   "Empresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Picture         =   "MDIUpdate.frx":27712
            Key             =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "MDIUpdate.frx":27FEC
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIUpdate.frx":288C6
            Key             =   "Plataforma"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Procesando"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Baltic"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBarEstado 
      Align           =   2  'Align Bottom
      Height          =   120
      Left            =   0
      TabIndex        =   1
      Top             =   3405
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   212
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CDialogDir 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu MArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MMantenimiento 
         Caption         =   "Actualizar"
         Shortcut        =   ^A
      End
      Begin VB.Menu MActOtraBase 
         Caption         =   "Actualicar otra Base"
         Shortcut        =   ^E
      End
      Begin VB.Menu Migracion 
         Caption         =   "Migración AT antes de julio"
      End
      Begin VB.Menu MOptimizarBaes 
         Caption         =   "Optimizar Base Dato"
      End
      Begin VB.Menu EliminarIdx 
         Caption         =   "Eliminar Indices"
      End
      Begin VB.Menu Reindexar 
         Caption         =   "Reindexar Tablas"
      End
      Begin VB.Menu MImportar 
         Caption         =   "Importar Bases Antiguas"
      End
      Begin VB.Menu MSalir 
         Caption         =   "Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MAmbiente 
      Caption         =   "Ambiente"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "MDIUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EliminarIdx_Click()
  Eliminar_Indices_SP
End Sub

Private Sub MActOtraBase_Click()
  FOtraBase.Show
End Sub

Private Sub MDIForm_Activate()
  MDI_Y_Max = MDIFormulario.ScaleHeight - 100
  MDI_X_Max = MDIFormulario.ScaleWidth - 100
 'MDIUpdate.Caption = "ACTUALIZACION A ENERO-2022"
End Sub

Private Sub MDIForm_Load()
Dim NumFile As Integer
Dim NumPos As Long
Dim RutaGeneraFile As String
Dim LineaTexto As String
Dim MiArchivo, MiRuta, MiNombre
Dim AnchoTemp As Single
Dim HayCnn As Boolean

  '---------------------------------------------------------------------------------
  'Datos de Conexion a la Base de Datos en las nubes mysql.diskcoversystem.com:13306
  '---------------------------------------------------------------------------------
   strBaseDatos = "diskcover_empresas"
   strServidor = "mysql.diskcoversystem.com"
   strUsuario = "diskcover"
   strPassword = "disk2017Cover"
   strPuerto = "13306"
   AdoStrCnnMySQL = "DRIVER={MySQL ODBC 5.1 Driver};" _
                  & "SERVER=" & strServidor & ";" _
                  & "DATABASE=" & strBaseDatos & ";" _
                  & "UID=" & strUsuario & ";" _
                  & "PASSWORD=" & strPassword & ";" _
                  & "PORT=" & strPuerto & ";"
  '---------------------------------------------------------------------------------
  Set MDIFormulario = Me
  AnchoPantalla = MDIFormulario.ScaleWidth - 100
  ConSubDir = False
  RutaSistema = Left$(CurDir$, 2) & "\SISTEMA"
  RutaSysBases = Left$(CurDir$, 2) & "\SYSBASES"
  RutaDestino = UCase$(Left$(CurDir$, 2))
  RutaSubDirTemp = RutaDestino
 ' FUnidad.Show 1
  RutaUpdate = RutaDestino & "\SISTEMA"
  RatonNormal
 'Cadena = InputBox("INGRESE LA UNIDAD DONDE ESTA EL SISTEMA", "UNIDAD DEL SISTEMA", Mid$(CurDir$, 1, 3))
  If RutaDestino = "" Then
     MsgBox "No Ingreso la ruta del Sistema"
     End
  Else
    'Leer_Lista_De_Impresoras
     RutaSistema = RutaDestino & "\SISTEMA"
     RutaSysBases = RutaDestino & "\SYSBASES"
     ChDir RutaSistema
    'UnidadSistema
     TipoModulo = Factu
     IngresarClave = True
    'MODULOS
     NumModulo = "98"
     Modulo = "UPDATE"
     MenuDeModulos = True
     TiempoSistema = Time
     Timer1.Interval = 10000
     IngresarClave = True
   
   RatonReloj
  'Determinar que tipo de bases utilizamos
   Evaluar = False
   SQL_Server = True
   Conectar_Base_Datos
   RutaEmpresa = UCase(RutaSistema & "\EMPRESA")
   RutaEmpresaOld = UCase(RutaSistema & "\EMPRESA")
 ' Verificamos si la base esta en Microsoft Access o en SQL Server 7.0
   FechaSistema = Format(date, FormatoFechas)
   NumEmpresa = "999"
   CodigoUsuario = "ACCESO99"
   NombreUsuario = "Actualizacion del Sistema"
   Empresa = "LA EMPRESA"
   EmailProcesos = "gerencia@diskcoversystem.com"
   Telefono1 = "09-9965-4196"
   Telefono2 = "09-8652-4396"
   Periodo_Contable = Ninguno
   LogoTipo = UCase(RutaSistema & "\LOGOS\DEFAULT.GIF")
   
    HayCnn = Get_WAN_IP
    Acceso_IP_PCs_SP_MySQL Si_No
  'Cambiamos a la ruta donde esta todo el sistema
'   RutaSistema = Unidad_Temp & "\SISTEMA"
'   RutaSysBases = Unidad_Temp & "\SYSBASES"
'   ChDir RutaSistema
   RatonNormal
  'Crear_Base_de_Datos()
  'MsgBox RutaEmpresa
  
   PonerDirEmpresa
   'NumEmpresa = "000"
   NumEmpresa = "."
   FActualizar.Show
   RatonNormal
  End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Control_Procesos "Q", "Salir Modulo de Update"
  End
End Sub

Private Sub Migracion_Click()
  If ClaveSupervisor Then
     RatonReloj
     Control_Procesos Normal, "Migración del 2007"
     FMigraAT.Show
  End If
End Sub

Private Sub MImportar_Click()
  Importar_Bases_Antiguas
End Sub

Private Sub MMantenimiento_Click()
  RatonReloj
  FActualizar.Show
End Sub

Private Sub MOptimizarBaes_Click()
   Optimizar_Memoria
End Sub

Private Sub MSalir_Click()
  Control_Procesos "Q", "Salir Modulo de Update"
  End
End Sub

Private Sub Reindexar_Click()
  Crear_Indices
  MsgBox "Proceso Terminado"
End Sub

Private Sub Timer1_Timer()
  Ver_Grafico_FormPict
End Sub

'''Public Sub Crear_Base_De_Datos()
'''Dim cat As ADOX.Catalog
'''Set cat = New ADOX.Catalog
'''' Crear la base de datos
'''cat.Create "Provider = Microsoft.Jet.OLEDB.4.0;" _
'''         & "Data Source = " & sNombreBase & ";"
'''End Sub
Private Sub Importar_Bases_Antiguas()
Dim SiID As Boolean
Dim SiItem As Boolean
Dim SiCod As Boolean

Dim ContTAB As Integer

Dim NumReg As Long
Dim TotalReg As Long

Dim CamposFile() As Campos_Tabla

Dim NombreTabla As String

    Progreso_Barra.Mensaje_Box = "SUBIENDO ABONOS DEL BANCO " & TextoBanco
    Progreso_Iniciar
    
    CDialogDir.Filename = RutaSysBases & "\Datos\Total\*.BDD"
    CDialogDir.InitDir = RutaSysBases & "\Datos\Total\"
    CDialogDir.Flags = cdlOFNFileMustExist + cdlOFNNoChangeDir + cdlOFNHideReadOnly
    CDialogDir.Filter = "Archivos BDD|*.BDD"
    CDialogDir.DialogTitle = "Abrir Archivo"
    CDialogDir.Action = 1
    J = InStrRev(CDialogDir.Filename, "\")
    If CDialogDir.Filename <> "" And J > 0 Then
       File1.Path = MidStrg(CDialogDir.Filename, 1, J)
       File1.Pattern = "*.BDD"
       Progreso_Barra.Incremento = 0
       For I = 0 To File1.ListCount - 1
           Progreso_Barra.Valor_Maximo = File1.ListCount
           NumReg = 1
           TotalReg = 2
           NumFile = FreeFile
           NombreArchivo = File1.Path & "\" & File1.List(I)
           Open NombreArchivo For Input As #NumFile
             Do While Not EOF(NumFile)
                Line Input #NumFile, Cod_Field
                Cod_Field = Replace(Cod_Field, vbCrLf, "")
                Select Case NumReg
                  Case 1
                       NombreTabla = TrimStrg(MidStrg(Cod_Field, InStrRev(Cod_Field, "-") + 1, Len(Cod_Field)))
                       Cod_Field = MidStrg(Cod_Field, 1, Len(Cod_Field) - Len(NombreTabla) - 2)
                       TotalReg = Val(TrimStrg(MidStrg(Cod_Field, InStrRev(Cod_Field, "-") + 1, Len(Cod_Field))))
                       Progreso_Barra.Mensaje_Box = NombreTabla
                       Progreso_Esperar
                       If Not Existe_Tabla(NombreTabla) Then GoTo Fin_Tabla
                  Case 2
                       SiID = False
                       SiItem = False
                       SiCod = False
                       ContTAB = 0
                       K = 1
                       For J = 1 To Len(Cod_Field)
                        If MidStrg(Cod_Field, J, 1) = vbTab Then
                           ReDim Preserve CamposFile(ContTAB) As Campos_Tabla
                           CamposFile(ContTAB).Campo = MidStrg(Cod_Field, K, J - K)
                           Select Case CamposFile(ContTAB).Campo
                             Case "ID": SiID = True
                             Case "Item": SiItem = True
                             Case "Codigo": SiCod = True
                           End Select
                           K = J + 1
                           ContTAB = ContTAB + 1
                        End If
                       Next J
                       Progreso_Barra.Mensaje_Box = "Encerando: " & NombreTabla
                       Progreso_Esperar
                       sSQL = "DELETE * " _
                            & "FROM " & NombreTabla & " "
                       If SiID Then
                          sSQL = sSQL & "WHERE ID > 0 "
                       ElseIf SiItem Then
                          sSQL = sSQL & "WHERE Item <> '.' "
                       ElseIf SiCod Then
                          sSQL = sSQL & "WHERE Codigo <> 'D' "
                       End If
                       Ejecutar_SQL_SP sSQL
'''                       Cadena = ""
'''                       For J = 0 To UBound(CamposFile)
'''                           Cadena = Cadena & CamposFile(J).Campo & " = " & CamposFile(J).Valor & vbCrLf
'''                       Next J
'''                       MsgBox Cadena
                  Case Else
                       ContTAB = 0
                       K = 1
                       For J = 1 To Len(Cod_Field)
                        If MidStrg(Cod_Field, J, 1) = vbTab Then
                           'MsgBox UBound(CamposFile) & vbCrLf & MidStrg(Cod_Field, K, J - K)
                           If ContTAB <= UBound(CamposFile) Then CamposFile(ContTAB).Valor = MidStrg(Cod_Field, K, J - K)
                           K = J + 1
                           ContTAB = ContTAB + 1
                        End If
                       Next J
                      'Insertamos el registro actual
'''                       Cadena = ""
                       SetAdoAddNew NombreTabla
                       For J = 0 To UBound(CamposFile)
                           If CamposFile(J).Campo <> "ID" Then SetAdoFields CamposFile(J).Campo, CamposFile(J).Valor
'''                           Cadena = Cadena & CamposFile(J).Campo & " = " & CamposFile(J).Valor & vbCrLf
                       Next J
                       SetAdoUpdate
'''                       MsgBox Cadena
                End Select
                Progreso_Barra.Mensaje_Box = NombreTabla & ": " & Format(NumReg, "#,##0") & " -> " & Format(TotalReg, "#,##0")
                Progreso_Esperar
                NumReg = NumReg + 1
                Contador = Contador + 1
             Loop
Fin_Tabla:
           Close #NumFile
       Next I
    End If
    Progreso_Final
End Sub



