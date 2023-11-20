VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form Respaldos 
   BackColor       =   &H00FF8080&
   Caption         =   "Espere un momento....     Estoy procesando las bases"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10965
   Icon            =   "Respaldo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   10965
   WindowState     =   2  'Maximized
   Begin VB.FileListBox File2 
      Height          =   675
      Left            =   8505
      TabIndex        =   20
      Top             =   5670
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   9870
      TabIndex        =   19
      Top             =   5670
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   9975
      TabIndex        =   18
      Top             =   6405
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox CheqCatalogos 
      BackColor       =   &H00FF8080&
      Caption         =   "Solo Catalogos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   8505
      TabIndex        =   17
      Top             =   1365
      Width           =   2220
   End
   Begin VB.CheckBox CheqSinRegNotas 
      BackColor       =   &H00FF8080&
      Caption         =   "Solo Educativo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   8505
      TabIndex        =   16
      Top             =   1050
      Width           =   2220
   End
   Begin VB.CheckBox CheqCatalogo 
      BackColor       =   &H00FF8080&
      Caption         =   "Subir con el Catalogo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   8505
      TabIndex        =   15
      Top             =   420
      Width           =   2220
   End
   Begin VB.CheckBox CheqFacturacion 
      BackColor       =   &H00FF8080&
      Caption         =   "Solo Facturacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   8505
      TabIndex        =   14
      Top             =   735
      Width           =   2220
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "DATOS DE LA EMPRESA A PROCESAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6210
      Left            =   1470
      TabIndex        =   11
      Top             =   105
      Width           =   6945
      Begin VB.ListBox LstArchivo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5910
         Left            =   105
         TabIndex        =   12
         Top             =   210
         Width           =   6735
      End
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   1050
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16761024
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
   Begin VB.CheckBox CheqAuditoria 
      BackColor       =   &H00FF8080&
      Caption         =   "Con Auditoria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   8505
      TabIndex        =   10
      Top             =   105
      Value           =   1  'Checked
      Width           =   2325
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
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
      Left            =   8505
      Picture         =   "Respaldo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4725
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&UnZip"
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
      Left            =   8505
      Picture         =   "Respaldo.frx":0E38
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2835
      Width           =   2325
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "S&ubir"
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
      Left            =   8505
      Picture         =   "Respaldo.frx":1142
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3780
      Width           =   2325
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Enviar"
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
      Left            =   8505
      Picture         =   "Respaldo.frx":19E8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1890
      Width           =   2325
   End
   Begin VB.ListBox LstTablas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   1680
      TabIndex        =   9
      Top             =   525
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.ListBox LGrupo 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   105
      TabIndex        =   1
      Top             =   1680
      Width           =   1275
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   1680
      Top             =   4410
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
      Left            =   1680
      Top             =   5355
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
      Left            =   1680
      Top             =   5040
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
   Begin MSAdodcLib.Adodc AdoOld 
      Height          =   330
      Left            =   1680
      Top             =   4725
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
      Caption         =   "Old"
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
   Begin MSAdodcLib.Adodc AdoAuxOld 
      Height          =   330
      Left            =   1680
      Top             =   5670
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
      Caption         =   "AuxOld"
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
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &HASTA:"
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
      TabIndex        =   13
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &GRUPO"
      BeginProperty Font 
         Name            =   "MS Serif"
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
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &DESDE:"
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
      TabIndex        =   2
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "Respaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AdoStrCnnOld As String
Dim AdoStrCnn1 As String
Dim AdoStrCnn2 As String
Dim PathEmpresa1 As String
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim XAdoStrCnn As String
Dim NombreEmpresa As String
Dim IJ As Long
Dim ITab As Long
Dim JCamp As Long
Dim Cont1 As Long
Dim ModuloResp As String
Dim Cod_FieldTabla As String
Dim NombTabla As String
Dim RetVale
Dim ProgBarr As Progreso_Barras

Public Sub Registro_De_Bases()
  Codigo = "F" & Format(Day(MBFechaI.Text), "00") & Format(Month(MBFechaI.Text), "00") & "000.TXT"
  RutaGeneraFile = Dir_Dialog.InitDir & Codigo
  NumFile = FreeFile
  FAConLineas = True
  Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
       Print #NumFile, TextoFileEmp
       Print #NumFile, String(55, "_")
  Close #NumFile
End Sub

Public Sub Empaquetar_Archivos_Zip()
Dim Resultado As Long
Dim intContadorFicheros As Integer
Dim FuncionesZip As ZIPUSERFUNCTIONS
Dim OpcionesZip As ZPOPT

FuncionesZip.DLLComment = DevolverDireccionMemoria(AddressOf FuncionParaProcesarComentarios)
FuncionesZip.DLLPassword = DevolverDireccionMemoria(AddressOf FuncionParaProcesarPassword)
FuncionesZip.DLLPrnt = DevolverDireccionMemoria(AddressOf FuncionParaProcesarMensajes)
FuncionesZip.DLLService = DevolverDireccionMemoria(AddressOf FuncionParaProcesarServicios)
'MsgBox NombreArchivoZip
  If NombreArchivoZip <> "" Then
     RatonReloj
     intContadorFicheros = 0
    'Determinamos la carpeta del respaldo
     Cadena = Dir(Dir_Dialog.InitDir & "*.txt", vbNormal)
     Do While Cadena <> ""
        If Cadena <> "." And Cadena <> ".." Then
           If (GetAttr(Dir_Dialog.InitDir & Cadena) And vbNormal) = vbNormal Then
              NombresFicherosZip.S(intContadorFicheros) = Dir_Dialog.InitDir & Cadena
              intContadorFicheros = intContadorFicheros + 1
              'MsgBox Cadena
           End If
        End If
        Cadena = Dir
     Loop
    'Nos cambiamos a la carpeta de respaldo
     ChDir Dir_Dialog.FinDir
    'Listamos los archivos a empaquetar
     Resultado = ZpInit(FuncionesZip)
     Resultado = ZpSetOptions(OpcionesZip)
     Resultado = ZpArchive(intContadorFicheros, NombreArchivoZip, NombresFicherosZip)
     RatonNormal
     MsgBox "SE GENERO SATISFACTORIAMENTE" & vbCrLf & vbCrLf _
          & "EL ARCHIVO DE RESPALDO EN: " & vbCrLf & vbCrLf _
          & NombreArchivoZip
     ChDir RutaEmpresa
  End If
  CurDir (Mid(RutaEmpresa, 1, 1))
  ChDir RutaEmpresa
End Sub

Public Sub Respaldo_Actual(Optional Es_Total As Boolean)
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim IdTime As Long
Dim strCnn As String
Dim TieneItem As Boolean
Dim TieneFecha As Boolean
Dim TienePeriodo As Boolean
Dim TieneWhere As Boolean
Dim OrdenarPor As String
Dim Progreso_Barra As Progreso_Barras

  LstTablas.Clear
  RatonReloj
  FAConLineas = True
  GrupoEmpresa = LGrupo.Text
  Contador = 0: FileResp = 0
 'Si Existe el archivo primero lo borramos
 'MsgBox NombreArchivoZip
  Cadena = Dir(NombreArchivoZip, vbArchive)
  If Cadena <> "" Then Kill NombreArchivoZip
  LstArchivo.AddItem "10 - Fecha_Respaldo"
  LstArchivo.AddItem "Datos_Respaldo:"
  LstArchivo.AddItem "01.- Empresa/Sucursal: " & Empresa
  LstArchivo.AddItem "02.- R.U.C.          : " & RUC
  LstArchivo.AddItem "03.- Gerente         : " & NombreGerente
  LstArchivo.AddItem "04.- Telefono        : " & Telefono1
  LstArchivo.AddItem "05.- FAX             : " & FAX
  LstArchivo.AddItem "06.- Numero Asignado : " & GrupoEmpresa
  LstArchivo.AddItem "07.- Fecha Inicial   : " & MBFechaI.Text & " (" & FechaDiaSem(MBFechaI.Text) & ")"
  LstArchivo.AddItem "08.- Fecha Final     : " & MBFechaF.Text & " (" & FechaDiaSem(MBFechaF.Text) & ")"
  LstArchivo.AddItem "09.- Modulo          : " & UCase(Modulo)
  LstArchivo.AddItem "10.- Usuario         : " & NombreUsuario
  LstArchivo.AddItem String(55, "_")
  LstArchivo.AddItem "ARCHIVOS PROCESADOS:"
  LstArchivo.AddItem "===================="
  LstArchivo.AddItem "F" & Format(Day(MBFechaI.Text), "00") & Format(Month(MBFechaI.Text), "00") & "000.TXT" & " => Fecha_Respaldo"
  TextoFileEmp = "10 - Fecha_Respaldo" & vbCrLf _
               & "Datos_Respaldo|" & vbCrLf _
               & "  1.- Empresa/Sucursal: " & Empresa & vbCrLf _
               & "  2.- R.U.C.          : " & RUC & vbCrLf _
               & "  3.- Gerente         : " & NombreGerente & vbCrLf _
               & "  4.- Telefono        : " & Telefono1 & vbCrLf _
               & "  5.- FAX             : " & FAX & vbCrLf _
               & "  6.- Numero Asignado : " & GrupoEmpresa & vbCrLf _
               & "  7.- Fecha Inicial   : " & MBFechaI.Text & " (" & FechaDiaSem(MBFechaI.Text) & ")" & vbCrLf _
               & "  8.- Fecha Final     : " & MBFechaF.Text & " (" & FechaDiaSem(MBFechaF.Text) & ")" & vbCrLf _
               & "  9.- Modulo          : " & UCase(Modulo) & vbCrLf _
               & " 10.- Usuario         : " & NombreUsuario & vbCrLf _
               & String(55, "_") & vbCrLf _
               & "ARCHIVOS PROCESADOS:" & vbCrLf _
               & "====================" & vbCrLf _
               & "F" & Format(Day(MBFechaI.Text), "00") & Format(Month(MBFechaI.Text), "00") & "000.TXT" & " => Fecha_Respaldo" & vbCrLf
' Eliminamos archivos de otros dias
  Cadena = Dir(Dir_Dialog.InitDir & "*.txt", vbArchive)
  If Cadena <> "" Then Kill Dir_Dialog.InitDir & "*.txt"
' Consultamos las cuentas de la tabla
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.Open AdoStrCnn
  Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
  Do Until RstSchema.EOF
     If RstSchema!TABLE_TYPE = "TABLE" And Mid(RstSchema!TABLE_NAME, 1, 1) <> "~" Then LstTablas.AddItem RstSchema!TABLE_NAME
     RstSchema.MoveNext
  Loop
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = LstTablas.ListCount + 1
  Progreso_Barra.Mensaje_Box = "PROGRESO DEL RESPALDO"
  Progreso_Esperar
  RatonReloj
  For ITab = 0 To LstTablas.ListCount - 1
      Progreso_Esperar
      Evaluar = False
      TieneItem = False
      TieneFecha = False
      TienePeriodo = False
      sSQL = "SELECT TOP 1 * " _
           & "FROM " & LstTablas.List(ITab) & " "
      SelectAdodc AdoQuery, sSQL
      With AdoQuery.Recordset
       For JCamp = 0 To .Fields.Count - 1
           If .Fields(JCamp).Name = "Item" Then
               TieneItem = True
           End If
           If .Fields(JCamp).Name = "Periodo" Then
               TienePeriodo = True
           End If
           If .Fields(JCamp).Name = "Fecha" Then
               TieneFecha = True
           End If
       Next JCamp
      End With
      If Es_Total Then TienePeriodo = False
      If LstTablas.List(ITab) = "Clientes_Facturacion" Then TienePeriodo = False
      If TieneItem Then Evaluar = True
      If CheqFacturacion.value Then
         Select Case LstTablas.List(ITab)
           Case "Catalogo_Bodegas", _
                "Catalogo_Lineas", _
                "Catalogo_Marcas", _
                "Catalogo_Productos", _
                "Clientes", _
                "Clientes_Datos_Extras", _
                "Clientes_Facturacion", _
                "Clientes_Matriculas", _
                "Detalle_Factura", _
                "Facturas", _
                "Trans_Abonos", _
                "Trans_Pedidos", _
                "Trans_Ticket"
                Evaluar = True
           Case Else: Evaluar = False
         End Select
      End If
      If CheqSinRegNotas.value Then
         Select Case LstTablas.List(ITab)
           Case "Catalogo_Cursos", _
                "Catalogo_Equivalencia", _
                "Catalogo_Estudiantil", _
                "Catalogo_Examen_Grado", _
                "Catalogo_Materias", _
                "Catalogo_Periodo_Lectivo", _
                "Clientes", _
                "Clientes_Datos_Extras", _
                "Clientes_Matriculas", _
                "Trans_Actas", _
                "Trans_Asistencia", _
                "Trans_Notas", _
                "Trans_Notas_Auxiliares", _
                "Trans_Notas_Grado", _
                "Trans_Promedios"
                Evaluar = True
           Case Else: Evaluar = False
         End Select
      End If
      If CheqCatalogos.value Then
         If Mid(LstTablas.List(ITab), 1, 8) = "Catalogo" Then
            Evaluar = True
         Else
            Evaluar = False
         End If
      End If
      If (CheqAuditoria.value <> 1) And (LstTablas.List(ITab) = "Trans_Entrada_Salida") Then Evaluar = False
      If Mid(LstTablas.List(ITab), 1, 4) = "Tipo" Then Evaluar = False
      If Mid(LstTablas.List(ITab), 1, 5) = "Tabla" Then Evaluar = False
      If Mid(LstTablas.List(ITab), 1, 5) = "Saldo" Then Evaluar = False
      If Mid(LstTablas.List(ITab), 1, 5) = "Fechas" Then Evaluar = False
      If Mid(LstTablas.List(ITab), 1, 7) = "Asiento" Then Evaluar = False
      If Mid(LstTablas.List(ITab), 1, 8) = "Balances" Then Evaluar = False
      'If Mid(LstTablas.List(ITab), 1, 8) = "Clientes" Then Evaluar = False
      If Evaluar Then
         FechaInicial = MBFechaI
         FechaFinal = MBFechaF
         OrdenarPor = ""
         TieneWhere = False
         sSQL = "SELECT * " _
              & "FROM " & LstTablas.List(ITab) & " "
         If TienePeriodo Then
            If TieneWhere Then
               sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
            Else
               sSQL = sSQL & "WHERE Periodo = '" & Periodo_Contable & "' "
               TieneWhere = True
            End If
            OrdenarPor = OrdenarPor & "Periodo,"
         End If
         If TieneItem Then
            If TieneWhere Then
               sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
            Else
               sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
               TieneWhere = True
            End If
            OrdenarPor = OrdenarPor & "Item,"
         End If
         If TieneFecha Then
            If TieneWhere Then
               sSQL = sSQL & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
            Else
               sSQL = sSQL & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
               TieneWhere = True
            End If
            OrdenarPor = OrdenarPor & "Fecha,"
          End If
         If OrdenarPor <> "" Then
            sSQL = sSQL & "ORDER BY " & Mid(OrdenarPor, 1, Len(OrdenarPor) - 1)
         End If
        'MsgBox sSQL
         SelectAdodc AdoQuery, sSQL
         With AdoQuery.Recordset
          If .RecordCount > 0 Then
              If TieneFecha Then
                 FechaInicial = .Fields("Fecha")
                .MoveLast
                 FechaFinal = .Fields("Fecha")
              End If
             .MoveFirst
              GenerarTablaArchivoPlano Respaldos, MBFechaI, FechaInicial, FechaFinal, LstTablas.List(ITab), AdoQuery
          End If
         End With
      End If
      LstArchivo.Refresh
  Next ITab
''
''  sSQL = "SELECT DISTINCT C.*,'CodDelDia' As CodigoAux " _
''       & "FROM Clientes As C " _
''       & "WHERE Fecha = #" & BuscarFecha(FechaSistema) & "# " _
''       & "AND Codigo <> '.' " _
''       & "UNION "
''  sSQL = sSQL _
''       & "SELECT DISTINCT C.*,Co.Codigo_B As CodigoAux " _
''       & "FROM Clientes As C,Comprobantes AS Co " _
''       & "WHERE Co.Item = '" & GrupoEmpresa & "' " _
''       & "AND C.Codigo = Co.Codigo_B " _
''       & "AND C.Codigo <> '.' " _
''       & "UNION "
''  sSQL = sSQL _
''       & "SELECT DISTINCT C.*,F.CodigoC As CodigoAux " _
''       & "FROM Clientes As C,Facturas As F " _
''       & "WHERE F.Item = '" & GrupoEmpresa & "' " _
''       & "AND C.Codigo = F.CodigoC " _
''       & "AND C.Codigo <> '.' " _
''       & "UNION "
''  sSQL = sSQL _
''       & "SELECT DISTINCT C.*,P.Cuenta_No As CodigoAux " _
''       & "FROM Clientes As C,Prestamos As P " _
''       & "WHERE P.Item = '" & GrupoEmpresa & "' " _
''       & "AND C.Codigo = P.Cuenta_No " _
''       & "AND C.Codigo <> '.' " _
''       & "UNION "
''  sSQL = sSQL _
''       & "SELECT DISTINCT C.*,CCP.Codigo As CodigoAux " _
''       & "FROM Clientes As C,Catalogo_CxCxP AS CCP " _
''       & "WHERE CCP.Item = '" & GrupoEmpresa & "' " _
''       & "AND C.Codigo = CCP.Codigo " _
''       & "AND C.Codigo <> '.' " _
''       & "ORDER BY C.Cliente "
''  SelectAdodc AdoQuery, sSQL
''  FechaInicial = MBFechaI
''  FechaFinal = MBFechaF
''  GenerarTablaArchivoPlano Respaldos, MBFechaI, FechaInicial, FechaFinal, "Clientes", AdoQuery, PictTabla
  LstArchivo.AddItem Codigo & " => Fecha_Respaldo"
  LstArchivo.Refresh
  Registro_De_Bases
  Respaldos.Caption = "MODULO DE RESPALDOS"
  RatonNormal
  Empaquetar_Archivos_Zip
End Sub

'''Public Sub RespaldoFamEnvios()
'''' Generamos Tabla:
'''  Respaldos.Caption = "Tabla: Remitentes..."
'''  If NumEmpresa = "001" Then
'''     sSQL = "SELECT DISTINCT R.* " _
'''          & "FROM Remitentes As R,Correos As Co " _
'''          & "WHERE Co.T = 'P' " _
'''          & "AND R.Codigo_R = Co.Cod_R " _
'''          & "ORDER BY Codigo_R "
'''  Else
'''     sSQL = "SELECT DISTINCT R.* " _
'''          & "FROM Remitentes As R,Correos As Co " _
'''          & "WHERE Co.Fecha_P BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''          & "AND Co.T = 'C' " _
'''          & "AND R.Codigo_R = Co.Cod_R " _
'''          & "ORDER BY Codigo_R "
'''  End If
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Remitentes", AdoQuery
'''' Generamos Tabla:
'''  Respaldos.Caption = "Tabla: Corresponsal..."
'''  sSQL = "SELECT * " _
'''       & "FROM Corresponsal " _
'''       & "WHERE Codigo_C <> '.' "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Corresponsal", AdoQuery
'''' Generamos Tabla:
'''  Respaldos.Caption = "Tabla: Corres_Envios..."
'''  sSQL = "SELECT * " _
'''       & "FROM Corres_Envios " _
'''       & "WHERE Codigo_C <> '.' "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Corres_Envios", AdoQuery
'''' Generamos Tabla:
'''  MiFecha = CLongFecha(CFechaLong(MBFechaI.Text) - 60)
'''  MiFecha = BuscarFecha(MiFecha)
'''
'''     If NumEmpresa = "001" Then
'''        sSQL = "SELECT C.*,Sucursal As SucIng,Sucursal As SucPag " _
'''             & "FROM Correos As C " _
'''             & "WHERE Sucursal <> '0' " _
'''             & "AND T = 'P' " _
'''             & "ORDER BY Envio_No "
'''        SelectData AdoQuery, sSQL
'''        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Correos", AdoQuery
'''        sSQL = "SELECT C.*,Sucursal As SucIng,Sucursal As SucPag " _
'''             & "FROM Correos As C " _
'''             & "WHERE Sucursal <> '0' " _
'''             & "AND C.Fecha BETWEEN #" & MiFecha & "# and #" & FechaFin & "# " _
'''             & "AND T = 'A' " _
'''             & "ORDER BY Envio_No "
'''        SelectData AdoQuery, sSQL
'''        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Correos", AdoQuery
'''     Else
'''        sSQL = "SELECT C.*,Sucursal As SucPag,Sucursal As SucIng " _
'''             & "FROM Correos As C " _
'''             & "WHERE Sucursal <> '0' " _
'''             & "AND C.T = 'C' " _
'''             & "AND C.Fecha_P BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''             & "ORDER BY C.Envio_No "
'''        SelectData AdoQuery, sSQL
'''        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Correos", AdoQuery
'''     End If
'''
'''  If NumEmpresa <> "001" Then
'''     sSQL = "SELECT '" & GrupoEmpresa & "' As Item,RLL.* " _
'''          & "FROM Resumen_Llamadas As RLL " _
'''          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''          & "ORDER BY Fecha,Envio_No "
'''     SelectData AdoQuery, sSQL
'''     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Resumen_Llamadas", AdoQuery
'''  End If
'''  RatonNormal
'''End Sub

Public Sub Respaldo_Access_97()
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Bancos..."
  sSQL = "SELECT * " _
       & "FROM Bancos " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Bancos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Catalogo_SubCtas..."
  sSQL = "SELECT (TC&TC&Codigo) As Codigo1,'" & NumEmpresa & "' As Item,TC,Beneficiario As Detalle,Presupuesto " _
       & "FROM Beneficiarios " _
       & "WHERE TC IN ('I','G') "
  SelectData AdoQuery, sSQL
  FechaInicial = MBFechaI
  FechaFinal = MBFechaF
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Catalogo_SubCtas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Catalogo_CxCxP..."
  sSQL = "SELECT (TC&TC&Codigo) As Codigo1,'" & GrupoEmpresa & "' As Item,TC,Cta " _
       & "FROM TransaccionesSC " _
       & "WHERE TC IN ('C','P','R') " _
       & "GROUP BY TC,Codigo,Cta "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Catalogo_CxCxP", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Catalogo_Cuentas..."
  sSQL = "SELECT '" & NumEmpresa & "' As Item,Ct.* " _
       & "FROM Catalogo As Ct " _
       & "WHERE TC <> 'X' " _
       & "ORDER BY Codigo "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Catalogo_Cuentas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Comprobantes..."
  sSQL = "SELECT C.*,C.CodigoB As Codigo_B " _
       & "FROM Comprobantes As C " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = " & Val(NumEmpresa) & " "
''  If CheqBT.Value = False Then
''     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
''     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
''     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
''     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
''     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
''  End If
 
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Comprobantes", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Conciliacion..."
  sSQL = "SELECT '" & NumEmpresa & "' As Item,C.* " _
       & "FROM Conciliacion As C " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Trans_Conciliacion", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Presupuestos..."
  sSQL = "SELECT '" & NumEmpresa & "' As Item,P.* " _
       & "FROM Presupuestos As P " _
       & "WHERE Cta <> '.' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Trans_Presupuestos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Retenciones..."
  sSQL = "SELECT '" & NumEmpresa & "' As Item,R.T,R.TP,R.Numero,R.Cta,R.Fecha," _
       & "(R.Ret_Porc/100) As Porc,R.Valor_Retenido As Valor_Ret,'303' As TD," _
       & "C.CodigoB As Codigo,'RF' As CodigoTR,Retencion As Retencion_No,R.Valor_Factura As Valor_Fact " _
       & "FROM Retenciones As R,Comprobantes As C " _
       & "WHERE R.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND R.TP = C.TP " _
       & "AND R.Numero = C.Numero " _
       & "AND R.Item = " & Val(NumEmpresa) & " " _
       & "ORDER BY R.TP,R.Numero,R.Cta,R.Fecha "
'''  If CheqBT.Value = False Then
'''     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
'''     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
'''     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
'''     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
'''     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
'''  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Trans_Retenciones", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Transacciones..."
  sSQL = "SELECT * " _
       & "FROM Transacciones " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = " & Val(NumEmpresa) & " "
'''  If CheqBT.Value = False Then
'''     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
'''     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
'''     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
'''     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
'''     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
'''  End If
' MsgBox sSQL
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Transacciones", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: TransaccionesGC..."
  sSQL = "SELECT * " _
       & "FROM TransaccionesGC " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = " & Val(NumEmpresa) & " "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Trans_Gastos_Caja", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: TransaccionesSC..."
  sSQL = "SELECT (TC&TC&Codigo) As Codigo1,C.* " _
       & "FROM TransaccionesSC As C " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = " & Val(NumEmpresa) & " "
'''  If CheqBT.Value = False Then
'''     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
'''     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
'''     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
'''     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
'''     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
'''  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Trans_SubCtas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Kardex..."
  sSQL = "SELECT 'PP' & Codigo_P As Codigo_P1,K.*,TP As TC,Cta As Contra_Cta " _
       & "FROM Kardex As K " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''  If CheqBT.Value = False Then
'''     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
'''     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
'''     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
'''     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
'''     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
'''  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, FechaInicial, FechaFinal, "Trans_Kardex", AdoQuery
''' Generamos Tabla:
''  Respaldos.Caption = "Tabla: Detalle_Factura..."
''  sSQL = "SELECT '" & NumEmpresa & "' As Item,('FA' & Codigo_C) As CodigoC,DF.*,'FA' As TC,Factura_No As Factura " _
''       & "FROM Detalle_Factura As DF " _
''       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
''  SelectData AdoQuery, sSQL
''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Detalle_Factura", AdoQuery, PictTabla
' Generamos Tabla:
'''  Respaldos.Caption = "Tabla: Facturas..."
'''  sSQL = "SELECT '" & NumEmpresa & "' As Item,('FA' & Codigo_C) As CodigoC,'FA' As TC,F.* " _
'''       & "FROM Facturas As F " _
'''       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Facturas", AdoQuery, PictTabla
'''' Generamos Tabla:
'''  Respaldos.Caption = "Tabla: Trans_Abonos..."
'''  sSQL = "SELECT T,'" & NumEmpresa & "' As Item,'FA' & Codigo_C As CodigoC,CtaxCob As Cta," _
'''       & "CtaxCob As Cta_CxP,TP,Fecha,Diario_No As Recibo_No,Factura,Abonos_MN As Abono " _
'''       & "FROM Diario_Caja " _
'''       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
'''       & "AND TP = 'CxC' "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Abonos", AdoQuery, PictTabla
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Suscripciones..."
'''     sSQL = "SELECT T,'" & NumEmpresa & "' As Item,'FA' & Codigo_C As Cuenta_No," _
'''          & "Area As Cta,TP,Contrato_No As Credito_No,S.Contador As Tasa," _
'''          & "S.Desde As Fecha,S.Hasta As Fecha_C " _
'''          & "FROM Contratos_Suscrip As S " _
'''          & "WHERE Contrato_No <> '.' "
'''     SelectData AdoQuery, sSQL
'''     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Prestamos", AdoQuery
' Generamos Tabla:
'''  Respaldos.Caption = "Tabla: Productos..."
'''  sSQL = "SELECT '" & NumEmpresa & "' As Item,Ct.*," _
'''       & "TP As TC,Cta_Inv As Cta_Inventario,Cta As Cta_Proveedor," _
'''       & "Cta1 As Cta_Costo_Venta,Cta_Ingreso As Cta_Ventas " _
'''       & "FROM Productos As Ct " _
'''       & "WHERE TP <> 'X' " _
'''       & "ORDER BY Codigo_Inv "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_Productos", AdoQuery, PictTabla
' Generamos Tabla:
'''  sSQL = "SELECT '" & NumEmpresa & "' As Item,Codigo As Codigo_Inv,Articulo As Producto," _
'''       & "'P' As TC,Cta_Inv As Cta_Inventario,Cta_CxP As Cta_Proveedor," _
'''       & "Cta_Costo As Cta_Costo_Venta,Cta_Ingreso As Cta_Ventas " _
'''       & "FROM Articulo As Ct " _
'''       & "WHERE CodigoL <> '.' " _
'''       & "ORDER BY Codigo "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_Productos", AdoQuery, PictTabla
' Generamos Tabla:
'''  Respaldos.Caption = "Tabla: Abono_De_Prestamo..."
'''  sSQL = "SELECT * " _
'''       & "FROM Abono_De_Prestamo " _
'''       & "WHERE DC <> 'X' "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Abono_De_Prestamo", AdoQuery, PictTabla
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Bloqueos..."
  sSQL = "SELECT * " _
       & "FROM Bloqueos " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = " & Val(NumEmpresa) & " "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Bloqueos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Cobranzas..."
  sSQL = "SELECT * " _
       & "FROM Cobranzas " _
       & "WHERE Porc_C <> 0 "
 ' SelectData AdoQuery, sSQL
 ' GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Tabla_Cobranzas", AdoQuery, PictTabla
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Conyugue..."
  sSQL = "SELECT *,Cuenta_No As Codigo " _
       & "FROM Conyugue " _
       & "WHERE Cuenta_No <> '.' "
'  SelectData AdoQuery, sSQL
'  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Clientes_Familiares", AdoQuery, PictTabla
' Generamos Tabla:
''  Respaldos.Caption = "Tabla: Cuentas..."
''  sSQL = "SELECT * " _
''       & "FROM Cuentas " _
''       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
''       & "AND Item = " & Val(NumEmpresa) & " "
''  SelectData AdoQuery, sSQL
''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Clientes_Libretas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Garantes..."
  sSQL = "SELECT * " _
       & "FROM Garantes " _
       & "WHERE TP <> 'X' "
'  SelectData AdoQuery, sSQL
'  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Garantes", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Monto_Apertura..."
  sSQL = "SELECT * " _
       & "FROM Monto_Apertura " _
       & "WHERE Monto_Aper<>0 "
'  SelectData AdoQuery, sSQL
'  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Monto_Apertura", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Prestamos..."
  sSQL = "SELECT '" & NumEmpresa & "' As Item,P.* " _
       & "FROM Prestamos As P " _
       & "WHERE P.Fecha >= #" & FechaIni & "# "
'  SelectData AdoQuery, sSQL
'  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Prestamos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Tasa_Interes..."
  sSQL = "SELECT * " _
       & "FROM Tasa_Interes " _
       & "WHERE Desde>=0 "
 ' SelectData AdoQuery, sSQL
  'GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Tabla_Interes", AdoQuery, PictTabla
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Tipo_Prestamo..."
  sSQL = "SELECT '" & NumEmpresa & "' As Item,TP.* " _
       & "FROM Tipo_Prestamo As TP " _
       & "WHERE TP <> '.' "
 ' SelectData AdoQuery, sSQL
 ' GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_Prestamo", AdoQuery, PictTabla
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Tipo_Proceso..."
  sSQL = "SELECT * " _
       & "FROM Tipo_Proceso " _
       & "WHERE DC <> '.' "
'  SelectData AdoQuery, sSQL
'  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_Proceso", AdoQuery, PictTabla
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Cajas..."
  sSQL = "SELECT * " _
       & "FROM Trans_Cajas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = " & Val(NumEmpresa) & " "
'  SelectData AdoQuery, sSQL
'  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Cajas", AdoQuery, PictTabla
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Libretas..."
  sSQL = "SELECT * " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = " & Val(NumEmpresa) & " "
 ' SelectData AdoQuery, sSQL
 ' GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Libretas", AdoQuery, PictTabla
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Prestamos..."
  sSQL = "SELECT '" & NumEmpresa & "' As Item,TP.* " _
       & "FROM Trans_Prestamos As TP " _
       & "WHERE TP.Fecha >= #" & FechaIni & "# "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Prestamos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Prestamos..."
  sSQL = "SELECT '" & NumEmpresa & "' As Item,TP.* " _
       & "FROM Trans_Prestamos As TP " _
       & "WHERE TP.Fecha_C " _
       & "BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND TP.T = 'C' "
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Prestamos", AdoQuery
End Sub

Public Sub Respaldo_Anterior()
  RatonReloj
  FAConLineas = True
  GrupoEmpresa = LGrupo.Text
  Contador = 0: FileResp = 0
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
 'Realizamos la Coneccion de la Base de Datos Access 97
  SQL_ServerOld = SQL_Server
  AdoStrCnn1 = AdoStrCnn
  SQL_Server = False
' Buscamos la cadena de conección a la base Sistema Antiguo
  RutaGeneraFile = RutaSistema & "\CONECTAR.TXT"
  NumFile = FreeFile
  AdoStrCnn2 = ""
  Open RutaGeneraFile For Input As #NumFile
    Do While Not EOF(NumFile)
       AdoStrCnn2 = AdoStrCnn2 & Input(1, #NumFile) ' Obtiene un carácter.
    Loop
  Close #NumFile
  If UCase(Mid(RutaDestino, Len(RutaDestino) - 2, 3)) = "MDB" Then
     AdoStrCnn2 = AdoStrCnn2 & "Data Source = " & RutaDestino     'PathEmpresa1
     AdoStrCnn = AdoStrCnn2
   ' Conexion con Sistema Antiguo
     ConectarAdodc AdoAuxOld
     ConectarAdodc AdoOld
    'Restauramos la coneccion Sistema Actual
     AdoStrCnn = AdoStrCnn1
    'Averiguamos la disponibilidad de la empresa
     Respaldos.Caption = "Creacion de la Empresa"
     NumEmpresa = "001"
     sSQL = "SELECT * " _
          & "FROM Empresas " _
          & "WHERE Item <> '.' " _
          & "ORDER BY Item DESC "
     SelectData AdoQuery, sSQL
     If AdoQuery.Recordset.RecordCount > 0 Then
        NumEmpresa = Format(Val(AdoQuery.Recordset.Fields("Item")) + 1, "000")
     End If
     NombreEmpresa = UCase("Empresa Numero " & NumEmpresa)
     SetAdoAddNew "Empresas"
     SetAdoFields "Grupo", NumEmpresa
     SetAdoFields "Item", NumEmpresa
     SetAdoFields "Empresa", NombreEmpresa
     SetAdoFields "SubDir", "EMPRE" & NumEmpresa
     SetAdoFields "Serie_Factura", "001001"
     SetAdoFields "Serie_Retencion", "001001"
     SetAdoFields "Autorizacion", "9999999999"
     SetAdoFields "AutorizacionR", "9999999999"
     SetAdoFields "CPais", "593"
     SetAdoFields "CProv", "00"
     SetAdoFields "Dec_PVP", 2
     SetAdoFields "Dec_Costo", 2
     SetAdoFields "Email", "diskcover@msn.com"
     SetAdoUpdate
     RatonReloj
   ' Preparamos los codigos de Clientes, Proveedores, Cuentas de Ahorro y Suscriptores
     Respaldos.Caption = "Insertando los Clientes nuevos"
     Preparar_Clientes_97
   ' Listo para migrar los datos de la base anterior
     AdoStrCnn = AdoStrCnn2
     MBFechaI = FechaSistema
     MBFechaF = FechaSistema
    'Comprobantes
     sSQL = "SELECT * " _
          & "FROM Comprobantes " _
          & "WHERE T <> 'X' " _
          & "ORDER BY Fecha "
     SelectAdodc AdoOld, sSQL
     With AdoOld.Recordset
      If .RecordCount > 0 Then
         .MoveLast
          MBFechaF = .Fields("Fecha")
          FechaIni = BuscarFecha(MBFechaI)
         .MoveFirst
          MBFechaI = .Fields("Fecha")
          FechaFin = BuscarFecha(MBFechaF)
      End If
     End With
     Respaldos.Caption = "Borrado de las Bases Actuales"
     AdoStrCnn = AdoStrCnn1
     sSQL = "DELETE * " _
          & "FROM Comprobantes " _
          & "WHERE Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Transacciones " _
          & "WHERE Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Trans_SubCtas " _
          & "WHERE Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Trans_Retenciones " _
          & "WHERE Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Catalogo_SubCtas " _
          & "WHERE Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
     sSQL = "DELETE * " _
          & "FROM Catalogo_CxCxP " _
          & "WHERE Item = '" & NumEmpresa & "' "
     ConectarAdoExecute sSQL
    'Comprobantes
     Contador = 0
     Respaldos.Caption = "Insertando Comprobantes"
     AdoStrCnn = AdoStrCnn2
     sSQL = "SELECT * " _
          & "FROM Comprobantes " _
          & "WHERE T <> 'X' " _
          & "ORDER BY Fecha "
     SelectAdodc AdoOld, sSQL
     With AdoOld.Recordset
      If .RecordCount > 0 Then
          AdoStrCnn = AdoStrCnn1
          Do While Not .EOF
             Contador = Contador + 1
             Respaldos.Caption = "Insertando Comprobantes: " & Format(Contador / .RecordCount, "00.0%")
             SetAdoAddNew "Comprobantes"
             SetAdoFields "T", .Fields("T")
             SetAdoFields "Codigo_B", .Fields("Autorizado")
             SetAdoFields "Concepto", .Fields("Concepto")
             SetAdoFields "Cotizacion", .Fields("Cotizacion")
             SetAdoFields "Efectivo", .Fields("Efectivo")
             SetAdoFields "TP", .Fields("TP")
             SetAdoFields "Numero", .Fields("Numero")
             SetAdoFields "Fecha", .Fields("Fecha")
             SetAdoFields "Item", NumEmpresa
             SetAdoUpdate
            .MoveNext
          Loop
      End If
     End With
     Contador = 0
    'Transacciones
     Respaldos.Caption = "Insertando Transacciones"
     AdoStrCnn = AdoStrCnn2
     sSQL = "SELECT * " _
          & "FROM Transacciones " _
          & "WHERE T <> 'X' " _
          & "ORDER BY Fecha "
     SelectAdodc AdoOld, sSQL
     With AdoOld.Recordset
      If .RecordCount > 0 Then
          AdoStrCnn = AdoStrCnn1
          ID_Trans = Maximo_De("Transacciones", "ID")
          Do While Not .EOF
             Contador = Contador + 1
             Respaldos.Caption = "Insertando Transacciones: " & Format(Contador / .RecordCount, "00.0%")
             SetAdoAddNew "Transacciones"
             SetAdoFields "T", .Fields("T")
             SetAdoFields "TP", .Fields("TP")
             SetAdoFields "Numero", .Fields("Numero")
             SetAdoFields "Fecha", .Fields("Fecha")
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "Cta", .Fields("Cta")
             SetAdoFields "Debe", .Fields("Debe")
             SetAdoFields "Haber", .Fields("Haber")
             SetAdoFields "Cheq_Dep", .Fields("Cheq_Dep")
             SetAdoFields "ID", ID_Trans
             SetAdoUpdate
             ID_Trans = ID_Trans + 1
            .MoveNext
          Loop
      End If
     End With
     Contador = 0
    'SubModulos
     Respaldos.Caption = "Insertando Transacciones de SubModulos"
     AdoStrCnn = AdoStrCnn2
     sSQL = "SELECT * " _
          & "FROM TransaccionesSC " _
          & "WHERE T <> 'X' " _
          & "ORDER BY Fecha "
     SelectAdodc AdoOld, sSQL
     With AdoOld.Recordset
      If .RecordCount > 0 Then
          AdoStrCnn = AdoStrCnn1
          ID_Trans = Maximo_De("Trans_SubCtas", "ID")
          Do While Not .EOF
             Contador = Contador + 1
             Respaldos.Caption = "Insertando Transacciones de SubModulos: " & Format(Contador / .RecordCount, "00.0%")
             SetAdoAddNew "Trans_SubCtas"
             SetAdoFields "T", .Fields("T")
             SetAdoFields "TP", .Fields("TP")
             SetAdoFields "Numero", .Fields("Numero")
             SetAdoFields "Fecha", .Fields("Fecha")
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "Fecha_V", .Fields("Fecha_V")
             SetAdoFields "Cta", .Fields("Cta")
             SetAdoFields "Codigo", .Fields("Codigo")
             SetAdoFields "Factura", .Fields("Factura")
             SetAdoFields "Debitos", .Fields("Debitos")
             SetAdoFields "Creditos", .Fields("Creditos")
             SetAdoFields "ID", ID_Trans
             SetAdoUpdate
             ID_Trans = ID_Trans + 1
            .MoveNext
          Loop
      End If
     End With
    'Catalogo de Cuentas
     Contador = 0
     Respaldos.Caption = "Insertando Catalogo de Cuentas"
     AdoStrCnn = AdoStrCnn2
     sSQL = "SELECT * " _
          & "FROM Catalogo " _
          & "WHERE Codigo <> '.' " _
          & "ORDER BY Codigo "
     SelectAdodc AdoOld, sSQL
     With AdoOld.Recordset
      If .RecordCount > 0 Then
          AdoStrCnn = AdoStrCnn1
          Do While Not .EOF
             Contador = Contador + 1
             Respaldos.Caption = "Insertando Catalogo de Cuentas: " & Format(Contador / .RecordCount, "00.0%")
             SetAdoAddNew "Catalogo_Cuentas"
             SetAdoFields "Clave", .Fields("Clave")
             SetAdoFields "TC", .Fields("TC")
             SetAdoFields "ME", .Fields("ME")
             SetAdoFields "DG", .Fields("DG")
             SetAdoFields "Codigo", .Fields("Codigo")
             SetAdoFields "Cuenta", .Fields("Cuenta")
             SetAdoFields "Presupuesto", .Fields("Presupuesto")
             SetAdoFields "Item", NumEmpresa
             SetAdoUpdate
            .MoveNext
          Loop
      End If
     End With
    'Catalogo de Cuentas
     Contador = 0
     Respaldos.Caption = "Insertando Catalogo de Cuentas"
     AdoStrCnn = AdoStrCnn2
     sSQL = "SELECT * " _
          & "FROM Beneficiarios " _
          & "WHERE Codigo <> '.' " _
          & "ORDER BY Codigo "
     SelectAdodc AdoOld, sSQL
     With AdoOld.Recordset
      If .RecordCount > 0 Then
          AdoStrCnn = AdoStrCnn1
          Do While Not .EOF
             Contador = Contador + 1
             Respaldos.Caption = "Insertando Catalogo de CxCxP: " & Format(Contador / .RecordCount, "00.0%")
             SetAdoAddNew "Catalogo_CxCxP"
             SetAdoFields "TC", .Fields("TC")
             SetAdoFields "Codigo", .Fields("Codigo")
             SetAdoFields "Item", NumEmpresa
             SetAdoUpdate
            .MoveNext
          Loop
      End If
     End With
     RatonNormal
     'Unload FEsperar
     MsgBox "La Empresa se Creo con el Nombre de: " & vbCrLf & vbCrLf & NombreEmpresa & " " & vbCrLf & vbCrLf _
            & "Vuelva a Ingresa al Sistema y verfique los datos Migrados"
     Unload Respaldos
  Else
     RatonNormal
     MsgBox "No se puede Migrar archivos que no sean de Microsoft Access"
  End If
End Sub

Public Sub AbrirCamposSQL(NumFile As Integer)
Dim Idx As Integer
Dim Cod_Base_Temp As String
    Cod_Emp = "": Cod_Base = "": Cod_Field = ""
    Line Input #NumFile, Cod_Base
     
    Cod_Base = Trim$(Replace(Cod_Base, "-", ""))
    
    Cod_Base_Temp = SinEspaciosIzq(Cod_Base)
    TotalReg = CLng(Cod_Base_Temp)
    Cod_Base = Trim(Mid(Cod_Base, Len(Cod_Base_Temp) + 1, Len(Cod_Base)))
    
    Cod_Base_Temp = SinEspaciosIzq(Cod_Base)
    FechaIni = BuscarFecha(Cod_Base_Temp)
    Cod_Base = Trim(Mid(Cod_Base, Len(Cod_Base_Temp) + 1, Len(Cod_Base)))
    
    Cod_Base_Temp = SinEspaciosIzq(Cod_Base)
    FechaFin = BuscarFecha(Cod_Base_Temp)
    Cod_Base = Trim(Mid(Cod_Base, Len(Cod_Base_Temp) + 1, Len(Cod_Base)))
    NombTabla = Cod_Base
     
    Line Input #NumFile, Cod_Field
    'MsgBox Cod_Base & vbCrLf & Cod_Field
    CantCampos = 0
    For Idx = 1 To Len(Cod_Field)
        If Mid(Cod_Field, Idx, 1) = "|" Then CantCampos = CantCampos + 1
    Next Idx
    ReDim TipoC(CantCampos + 1) As Campos_Tabla
    No_Desde = 1: No_Hasta = 1
    Cadena = Cod_Field
    For Idx = 1 To CantCampos
        Do
           No_Hasta = No_Hasta + 1
        Loop Until Mid(Cadena, No_Hasta, 1) = "|"
        TipoC(Idx).Campo = Trim(Mid(Cadena, No_Desde, No_Hasta - 1))
        Cadena = Mid(Cadena, No_Hasta + 1, Len(Cadena))
        No_Desde = 1: No_Hasta = 1
    Next Idx
End Sub

Public Sub LeerTablasPlanas(NombreFile As String)
  If Mid(NombreFile, 1, 3) <> "ACT" Then
     NumFile = FreeFile
     sSQL = "SELECT * FROM " & Cod_Base & " " _
          & "WHERE Item <> 255 "
     SelectAdodc AdoQuery, sSQL
     RutaGeneraFile = RutaSysBases & "\" & NombreFile
     Open RutaGeneraFile For Input As #NumFile
     Line Input #NumFile, Cadena
     Cod_Base = Cadena
     Line Input #NumFile, Cadena
     Cod_Field = Cadena
     Do While Not EOF(NumFile)
        Line Input #NumFile, Cadena
     Loop
     Close #NumFile
  End If
End Sub

'''Public Sub EjecutarRespaldo()
'''  RatonReloj
'''  File1.Filename = Dir1.Path & "\*.ZIP"
'''  File2.Filename = Dir1.Path & "\*.TXT"
'''  If File2.ListCount > 0 Then Kill RutaBackup & "\*.TXT"
'''  TextUnidad.Text = UCase(TextUnidad.Text)
'''  Respaldos.Caption = "Restaurando las bases "
'''  If (TextUnidad.Text & ":" = "A:") Or (TextUnidad.Text & ":" = "B:") Then
'''     ChDrive TextUnidad.Text & ":"
'''     Shell "restaura.bat " & TextUnidad.Text & ": " & Mid(RutaSistema, 1, 2) & " " & File1.Filename, vbMaximizedFocus
'''  Else
'''     ChDrive TextUnidad.Text & ":"
'''     ChDir TextUnidad.Text & ":\SYSBASES"
'''     Shell "restaura.bat " & TextUnidad.Text & ": Ninguno " & File1.Filename, vbMaximizedFocus
'''  End If
'''  ChDrive Mid(RutaSistema, 1, 2)
'''  Respaldos.Caption = "RESPALDOS DE BASES"
'''  File1.Filename = Dir1.Path & "\*.ZIP"
'''  File2.Filename = Dir1.Path & "\*.TXT"
'''  RatonNormal
'''End Sub

Public Sub AbrirArchivoSQL(NumFile As Integer)
    RatonReloj
    LstArchivo.Clear
    Cod_Emp = ""
    Cod_Base = ""
    Cod_Field = ""
    Cod_NumEmp = ""
    Cod_FechaI = ""
    Cod_FechaF = ""
    Do While Not EOF(NumFile)
       Line Input #NumFile, Cadena
       LstArchivo.AddItem Cadena
       Select Case Val(Mid(Cadena, 1, 3))
         Case 1: Cod_Emp = Mid(Cadena, 25, Len(Cadena) - 25)
         Case 6: Cod_NumEmp = Format(Mid(Cadena, 25, 3), "000")
         Case 7: Cod_FechaI = Mid(Cadena, 25, 10)
         Case 8: Cod_FechaF = Mid(Cadena, 25, 10)
       End Select
    Loop
    Close #NumFile
    RatonNormal
End Sub

Public Sub ActualizarRangoFecha(NumFile As Integer, _
                                NombreTabla As String)
Dim Cont1 As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim SiEncontro As Boolean
Dim TienePeriodo As Boolean
Dim JCamp As Integer
  RatonReloj
  ID_Trans = -1
  TienePeriodo = False
'''  FechaIni = BuscarFecha(MBFechaI)
'''  FechaFin = BuscarFecha(MBFechaF)
  NombreTabla = Trim(NombreTabla)
  Cont1 = 0
  
  sSQL = "SELECT * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE Item = '" & NumEmpresa & "' "
  SelectAdodc AdoQuery, sSQL
  With AdoQuery.Recordset
   For JCamp = 0 To .Fields.Count - 1
       If .Fields(JCamp).Name = "Periodo" Then TienePeriodo = True
       If .Fields(JCamp).Name = "ID" Then ID_Trans = 0
   Next JCamp
  End With
  
  sSQL = "DELETE * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  If TienePeriodo Then sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
  ConectarAdoExecute sSQL
  
  sSQL = "SELECT * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE Item = '" & NumEmpresa & "' "
  If TienePeriodo Then sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
  If ID_Trans >= 0 Then sSQL = sSQL & "ORDER BY ID DESC "
  'MsgBox sSQL & vbCrLf & TienePeriodo & vbCrLf & ID_Trans
  SelectAdodc AdoQuery, sSQL
  If AdoQuery.Recordset.RecordCount > 0 Then
     If ID_Trans >= 0 Then ID_Trans = AdoQuery.Recordset.Fields("ID")
  End If
  ID_Trans = ID_Trans + 1
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     Progreso_Barra.Mensaje_Box = "[" & Cont1 & "] " & NombreTabla
     Progreso_Esperar True

     'MsgBox "Fin "
     LeerCamposTabla Cod_FieldTabla, Ninguno
     SetCamposTabla True
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarFecha(BCodigo As Variant, _
                           NumFile As Integer, _
                           NombreTabla As String)
Dim Contl As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  NumItemTemp = NumEmpresa
  Cont1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     Si_No = True
     Progreso_Barra.Mensaje_Box = "[" & Contl & "] " & NombreTabla
     Progreso_Esperar True

     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     sSQL = "SELECT * FROM " & NombreTabla & " " _
          & "WHERE " & BCodigo & " = '" & CodBusq & "' " _
          & "AND Fecha = #" & Mifecha & "# "
     SelectAdodc AdoQuery, sSQL
     If AdoQuery.Recordset.RecordCount <= 0 Then
        SetCamposTabla True
     End If
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarMayor(NumFile As Integer, _
                           NombreTabla As String)
Dim Contl As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim SiEncontro As Boolean
  RatonReloj
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  NombreTabla = Trim(NombreTabla)
  Cont1 = 0
  sSQL = "DELETE * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE Fecha >= #" & FechaIni & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  ConectarAdoExecute sSQL
  sSQL = "SELECT * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE Item = '" & NumEmpresa & "' "
  SelectAdodc AdoQuery, sSQL
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     Progreso_Barra.Mensaje_Box = "[" & Contl & "] " & NombreTabla
     Progreso_Esperar True

     LeerCamposTabla Cod_FieldTabla, Ninguno
     SetCamposTabla True
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarTablaCompletaItem(NombreTabla As String)
Dim Cont1 As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim SiEncontro As Boolean
  RatonReloj
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  NombreTabla = Trim(NombreTabla)
  Cont1 = 0
  sSQL = "DELETE * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE Item = '" & NumEmpresa & "' "
  ConectarAdoExecute sSQL
  sSQL = "SELECT * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE Item = '" & NumEmpresa & "' "
  SelectAdodc AdoQuery, sSQL
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     Progreso_Barra.Mensaje_Box = "[" & Cont1 & "] " & NombreTabla
     Progreso_Esperar True

     LeerCamposTabla Cod_FieldTabla, Ninguno
     SetCamposTabla True
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarCodigo(BCodigo As Variant, _
                            NumFile As Integer, _
                            NombreTabla As String)
Dim Contl As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  NombreTabla = Trim(NombreTabla)
  BCodigo = Trim(BCodigo)
  sSQL = "SELECT * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE " & BCodigo & " <> '.' " _
       & "ORDER BY " & BCodigo & " "
  SelectAdodc AdoQuery, sSQL
  Cont1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     Progreso_Barra.Mensaje_Box = "[" & Contl & "] " & NombreTabla
     Progreso_Esperar True
     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     With AdoQuery.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
          .Find (BCodigo & " = '" & CodBusq & "' ")
       End If
       If Not .EOF Then SetCamposTabla False Else SetCamposTabla True
     End With
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarCodigoN(BCodigo As Variant, _
                             NumFile As Integer, _
                             NombreTabla As String)
Dim Contl As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  sSQL = "SELECT * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE " & BCodigo & " <> 0 " _
       & "ORDER BY " & BCodigo & " "
  SelectAdodc AdoQuery, sSQL
  Cont1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = 0
     Progreso_Barra.Mensaje_Box = "[" & Contl & "] " & NombreTabla
     Progreso_Esperar True

     CodBusq = Val(LeerCamposTabla(Cod_FieldTabla, BCodigo))
     With AdoQuery.Recordset
         .MoveFirst
         .Find (BCodigo & " = " & CodBusq & " ")
          If Not .EOF Then
             SetCamposTabla False
          Else
             SetCamposTabla True
          End If
     End With
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarCodigoP(BCodigo As Variant, _
                             NumFile As Integer, _
                             NombreTabla As String)
Dim Contl As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  NumItemTemp = NumEmpresa
  Cont1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     Si_No = True
     Progreso_Barra.Mensaje_Box = "[" & Contl & "] " & NombreTabla
     Progreso_Esperar True

     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     sSQL = "SELECT * " _
          & "FROM " & NombreTabla & " " _
          & "WHERE Credito_No = '" & Contrato_No & "' " _
          & "AND " & BCodigo & " = '" & CodBusq & "' " _
          & "AND Fecha = #" & Mifecha & "# " _
          & "AND Mes_No = " & NoMeses & " " _
          & "AND TP = '" & TipoProc & "' " _
          & "AND Item = '" & NumItemTemp & "' "
     SelectAdodc AdoQuery, sSQL
     If AdoQuery.Recordset.RecordCount > 0 Then
        SetCamposTabla False
     End If
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarCodigoCta(BCodigo As Variant, _
                               NumFile As Integer, _
                               NombreTabla As String)
Dim Contl As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  NumItemTemp = NumEmpresa
  Cont1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     Si_No = True
     Progreso_Barra.Mensaje_Box = "[" & Contl & "] " & NombreTabla
     Progreso_Esperar True

     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     sSQL = "SELECT * " _
          & "FROM " & NombreTabla & " " _
          & "WHERE Cta = '" & Cta & "' " _
          & "AND " & BCodigo & " = '" & CodBusq & "' " _
          & "AND Item = '" & NumItemTemp & "' "
     If Empleados Then sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
     SelectAdodc AdoQuery, sSQL
     If AdoQuery.Recordset.RecordCount > 0 Then
        SetCamposTabla False
     Else
        SetCamposTabla True
     End If
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarItem(NumFile As Integer, _
                          NombreTabla As String, _
                          BCodigo As Variant)
Dim Contl As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  NumItemTemp = NumEmpresa
  MsgBox NumEmpresa
  sSQL = "DELETE * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE Item = '" & NumItemTemp & "' "
  ConectarAdoExecute sSQL
  Cont1 = 0
 'If NombreTabla = "Catalogo_Rol_Rubros" Then MsgBox "..."
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     'MsgBox Cod_FieldTabla
     CodBusq = Ninguno
     Si_No = True
     Empleados = False
     Progreso_Barra.Mensaje_Box = "[" & Contl & "] " & NombreTabla
     Progreso_Esperar True

     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     sSQL = "SELECT * " _
          & "FROM " & NombreTabla & " " _
          & "WHERE Item = '" & NumItemTemp & "' "
     SelectAdodc AdoQuery, sSQL
     SetCamposTabla True
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarCodigoItem(BCodigo As Variant, _
                                NumFile As Integer, _
                                NombreTabla As String)
Dim Contl As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  NumItemTemp = NumEmpresa
  'MsgBox CheqCatalogo.value
  If CheqCatalogo.value = 1 And NombreTabla = "Catalogo_Cuentas" Then
     sSQL = "DELETE * " _
          & "FROM " & NombreTabla & " " _
          & "WHERE Item = '" & NumItemTemp & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
     ConectarAdoExecute sSQL
  End If
  Cont1 = 0
 'If NombreTabla = "Catalogo_Rol_Rubros" Then MsgBox "..."
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     'MsgBox Cod_FieldTabla
     CodBusq = Ninguno
     Si_No = True
     Empleados = False
     Progreso_Barra.Mensaje_Box = "[" & Cont1 & "] " & NombreTabla
     Progreso_Esperar True

     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     
     sSQL = "SELECT * " _
          & "FROM " & NombreTabla & " " _
          & "WHERE " & BCodigo & " = '" & CodBusq & "' " _
          & "AND Item = '" & NumItemTemp & "' "
     If Empleados Then sSQL = sSQL & "AND Periodo = '" & Periodo_Contable & "' "
     SelectAdodc AdoQuery, sSQL
     If AdoQuery.Recordset.RecordCount > 0 Then
        SetCamposTabla False
     Else
        SetCamposTabla True
     End If
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub Actualizar_Bases_DiskCover()
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim IdTime As Long
Dim strCnn As String
  LstTablas.Clear
  LstArchivo.Clear
  LstArchivo.AddItem "TABLAS PROCESADAS CON DATOS:"
  RatonReloj
  GrupoEmpresa = LGrupo.Text
  Contador = 0: FileResp = 0
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
' Eliminamos archivos de otros dias
  PathEmpresa1 = Dir1.Path & "\" & File1.Filename
  File1.Filename = Dir1.Path & "\F*.TXT"
  File1.Refresh
' Consultamos las cuentas de la tabla
  Set AdoCon1 = New ADODB.Connection
  AdoCon1.Open AdoStrCnn
  Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
  Do Until RstSchema.EOF
     If RstSchema!TABLE_TYPE = "TABLE" And Mid(RstSchema!TABLE_NAME, 1, 1) <> "~" Then LstTablas.AddItem RstSchema!TABLE_NAME
     RstSchema.MoveNext
  Loop
  RatonReloj
  'ProgBarra.value = 0: ProgBarra.Max = LstTablas.ListCount + 5
  For ITab = 0 To LstTablas.ListCount - 1
      'Progreso_Esperar True
      Progreso_Esperar True
      LstArchivo.AddItem "   - " & LstTablas.List(ITab)
      Si_No = False
      Encontro = False
      sSQL = "SELECT * " _
           & "FROM " & LstTablas.List(ITab) & " "
      SelectAdodc AdoQuery, sSQL
      With AdoQuery.Recordset
       For JCamp = 0 To .Fields.Count - 1
           If .Fields(JCamp).Name = "Item" Then Si_No = True
           If .Fields(JCamp).Name = "Fecha" Then Encontro = True
       Next JCamp
      End With
      If Si_No Then
         Evaluar = True
         If Mid(LstTablas.List(ITab), 1, 4) = "Tipo" Then Evaluar = False
         If Mid(LstTablas.List(ITab), 1, 5) = "Tabla" Then Evaluar = False
         If Mid(LstTablas.List(ITab), 1, 5) = "Saldo" Then Evaluar = False
         If Mid(LstTablas.List(ITab), 1, 5) = "Fechas" Then Evaluar = False
         If Mid(LstTablas.List(ITab), 1, 7) = "Asiento" Then Evaluar = False
         If Mid(LstTablas.List(ITab), 1, 8) = "Balances" Then Evaluar = False
         If Evaluar Then
            For IJ = 0 To File1.ListCount - 1
                RutaGeneraFile = RutaSysBases & "\DATOS\R" & GrupoEmpresa & "\" & File1.List(IJ)
                NumFile = FreeFile
                Open RutaGeneraFile For Input As #NumFile
                     AbrirCamposSQL NumFile
                     If LstTablas.List(ITab) = Cod_Base Then LstArchivo.AddItem String(25 - Len(LstTablas.List(ITab)), ".") & "Ok"
                     Do While Not EOF(NumFile)
                        Line Input #NumFile, Cod_FieldTabla
                        'CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
                        ' SetCamposTabla False
                     Loop
                Close #NumFile
            Next IJ
         End If
      End If
  Next ITab
  'ProgBarra.value = ProgBarra.Max
'''
'''
'''Dim Contl As Long
'''Dim Cod_FieldTabla As String
'''Dim Valor_Field As String
'''Dim CodBusq As Variant
'''Dim SiEncontro As Boolean
'''  RatonReloj
'''  BCodigo = Trim(BCodigo)
'''  NombreTabla = Trim(NombreTabla)
'''  NumItemTemp = NumEmpresa
'''  Cont1 = 0
'''  Do While Not EOF(NumFile)
'''     Line Input #NumFile, Cod_FieldTabla
'''     CodBusq = Ninguno
'''     Si_No = True
'''     Progreso_Esperar True
'''     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
'''     sSQL = "SELECT * FROM " & NombreTabla & " " _
'''          & "WHERE " & BCodigo & " = '" & CodBusq & "' " _
'''          & "AND Item = '" & NumItemTemp & "' "
'''     SelectAdodc AdoQuery, sSQL
'''     If AdoQuery.Recordset.RecordCount > 0 Then
'''        SetCamposTabla False
'''     Else
'''        SetCamposTabla True
'''     End If
'''     Cont1 = Cont1 + 1
'''  Loop
  RatonNormal
End Sub

Public Sub ActualizarCodigoC(BCodigo As Variant, _
                             NumFile As Integer, _
                             NombreTabla As String)
Dim Contl As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
Dim SiExist As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  sSQL = "SELECT * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE " & BCodigo & " <> '.' " _
       & "ORDER BY " & BCodigo & " "
  SelectAdodc AdoQuery, sSQL
  Cont1 = 0
  Do While Not EOF(NumFile)
     SiExist = True
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     NumItemTemp = NumEmpresa
     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     Progreso_Barra.Mensaje_Box = "[" & Contl & "] " & NombreTabla
     Progreso_Esperar True

     With AdoQuery.Recordset
         .MoveFirst
         .Find (BCodigo & " = '" & CodBusq & "' ")
          If Not .EOF Then
             If NumEmpresa = "001" And .Fields("T") = "C" Then
                SetCamposTabla False
             End If
          Else
             SetCamposTabla True
          End If
     End With
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Private Sub Command1_Click()
Dim RutaRespaldos As String
  RutaOrigen = RutaSysBases & "\DATOS\D" & LGrupo & "\"
  RutaRespaldos = RutaOrigen
  Dir_Dialog.Filename = RutaOrigen & "*.Zip"
  Dir_Dialog.Filter = "Archivos Zip|*.zip"
  Dir_Dialog.InitDir = RutaOrigen & "*.zip"
  Dir_Dialog.Filename = Abrir_Archivo(Me.hwnd, Dir_Dialog, OpenZip)
 'MsgBox Dir_Dialog.File
  RutaOrigen = UCase(Dir_Dialog.Filename)
  NombreArchivoZip = RutaOrigen
 'MsgBox NombreArchivoZip
  If RutaOrigen <> "" Then
     If Mid(Dir_Dialog.File, 1, 2) = "DS" Then Si_No = True Else Si_No = False
     Codigo1 = Trim(Mid(Dir_Dialog.File, 3, 3))
     Codigo2 = Trim(Mid(Dir_Dialog.File, Len(Dir_Dialog.File) - 3, 3))
    'MsgBox RutaOrigen & vbCrLf & Codigo1 & vbCrLf & Codigo2 & vbCrLf & Si_No
     If (Codigo1 = LGrupo) And IsNumeric(Codigo1) And UCase(Codigo2) = "ZIP" And Si_No Then
        RatonReloj
        LGrupo = Codigo1
        NumEmpresa = LGrupo
        GrupoEmpresa = LGrupo
        ConSubDir = False
        Contador = 0: FileResp = 0
        FechaValida MBFechaI
        FechaValida MBFechaF
      ' Eliminamos archivos de otros dias
        Cadena = Dir(RutaRespaldos & "*.txt", vbNormal)
        If Cadena <> "" Then Kill RutaRespaldos & "*.txt"
       'Pasamos a descomprimir
        UnZip RutaOrigen, RutaRespaldos
        MsgBox "FIN DEL PROCESO DE DESCOMPRESION," & vbCrLf & vbCrLf _
             & "PROCEDA A SUBIR LA INFORMACION"
     ElseIf ConSucursal And IsNumeric(Codigo1) And UCase(Codigo2) = "ZIP" And Si_No Then
        RatonReloj
        LGrupo = Codigo1
        NumEmpresa = Codigo1
        GrupoEmpresa = Codigo1
        ConSubDir = False
        Contador = 0: FileResp = 0
        FechaValida MBFechaI
        FechaValida MBFechaF
      ' Eliminamos archivos de otros dias
        Cadena = Dir(RutaRespaldos & "*.txt", vbNormal)
        If Cadena <> "" Then Kill RutaRespaldos & "*.txt"
       'Pasamos a descomprimir
        UnZip RutaOrigen, RutaRespaldos
        MsgBox "FIN DEL PROCESO DE DESCOMPRESION," & vbCrLf & vbCrLf _
             & "PROCEDA A SUBIR LA INFORMACION " & NumEmpresa
     Else
        MsgBox "ESTE ARCHIVO NO ES VALIDO," & vbCrLf & vbCrLf _
             & "NO SE PUEDE PROCESAR"
     End If
  End If
  Respaldos.Caption = "MODULO DE RESPALDOS"
  RatonNormal
  MBFechaI.SetFocus
End Sub

Private Sub Command2_Click()
  Unload Respaldos
End Sub

''Private Sub Command3_Click()
''  Actualizar_Bases_DiskCover
''End Sub

Private Sub Command4_Click()
  LstArchivo.Clear
  sSQL = "UPDATE Accesos " _
       & "SET Item = '" & NumEmpresa & "' "
  'ConectarAdoExecute sSQL
  ConSubDir = False
  If NumEmpresa = "" Then NumEmpresa = LGrupo.Text
  RutaOrigen = RutaBackup & "\DATOS\D" & LGrupo.Text
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  Codigo = "DS" & GrupoEmpresa & "-" & FechaDeZip(MBFechaI)
  If MBFechaI <> MBFechaF Then Codigo = Codigo & "-" & FechaDeZip(MBFechaF)
  Dir_Dialog.Filename = RutaOrigen & "\" & Codigo & ".Zip"
  Dir_Dialog.Filter = "Archivos Zip|*.zip"
  Dir_Dialog.InitDir = RutaOrigen & "\"
  Dir_Dialog.Filename = Abrir_Archivo(Me.hwnd, Dir_Dialog, SaveFile)
  RutaDestino = UCase(Dir_Dialog.Filename)
  NombreArchivoZip = RutaDestino
  If RutaDestino <> "" Then Respaldo_Actual
  Unload Respaldos
End Sub


Private Sub Command5_Click()
Dim AuxNumEmp As String
Dim Progreso_Barra As Progreso_Barras
Dim RutaRespaldos As String
Dim ArchivosRespaldos() As String

  AuxNumEmp = NumEmpresa
  NumEmpresa = Trim(Cod_NumEmp)
  If NumEmpresa = "" Then NumEmpresa = LGrupo.Text
  RutaRespaldos = RutaSysBases & "\DATOS\D" & LGrupo.Text & "\"
  TotalReg = 0
  Cadena = Dir(RutaRespaldos & "*.txt", vbNormal)
  Do While Cadena <> ""
     If Cadena <> "." And Cadena <> ".." Then
        If (GetAttr(RutaRespaldos & Cadena) And vbNormal) = vbNormal Then
           TotalReg = TotalReg + 1
        End If
     End If
     Cadena = Dir
  Loop
  ReDim ArchivosRespaldos(TotalReg) As String
  TotalReg = 0
  Cadena = Dir(RutaRespaldos & "*.txt", vbNormal)
  Do While Cadena <> ""
     If Cadena <> "." And Cadena <> ".." Then
        If (GetAttr(RutaRespaldos & Cadena) And vbNormal) = vbNormal Then
           ArchivosRespaldos(TotalReg) = RutaRespaldos & Cadena
           TotalReg = TotalReg + 1
        End If
     End If
     Cadena = Dir
  Loop
  Contador = 0: FileResp = 0
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  ConectarAdodc AdoAct
  ConectarAdodc AdoAux
  ConectarAdodc AdoQuery
  Progreso_Barra.Incremento = 0
  Progreso_Barra.Valor_Maximo = TotalReg + 10
  Progreso_Barra.Mensaje_Box = "PROGRESO DE LA RESTAURACION"
  Progreso_Esperar
  LstArchivo.Clear
  LstArchivo.AddItem "PROCESO DE ACTUALIZACION DE DATOS:"
    
    For IJ = 0 To UBound(ArchivosRespaldos) - 1
        RatonReloj
        Progreso_Esperar
       'MsgBox ArchivosRespaldos(IJ)
        RutaGeneraFile = ArchivosRespaldos(IJ)
        NombreArchivo = ArchivosRespaldos(IJ)
       'MsgBox RutaGeneraFile
        NumFile = FreeFile
        Open RutaGeneraFile For Input As #NumFile
             AbrirCamposSQL NumFile
             ProgBarr.Incremento = 0
             ProgBarr.Valor_Maximo = TotalReg
             ProgBarr.Mensaje_Box = NombTabla
             'Pict_Proceso PictTabla, ProgBarr
             
             LstArchivo.AddItem "Procesando el archivo " & NombreArchivo
             LstArchivo.AddItem "De la tabla: " & NombTabla
            'MsgBox Cod_Base
             If Cod_Base = "Fecha_Respaldo" Then
                AbrirArchivoSQL NumFile
                MBFechaI = Cod_FechaI
                MBFechaF = Cod_FechaF
                FechaValida MBFechaI
                FechaValida MBFechaF
                FechaIni = BuscarFecha(MBFechaI)
                FechaFin = BuscarFecha(MBFechaF)
             End If
            'MsgBox FechaIni & vbCrLf & FechaFin & vbCrLf & Cod_Base
             Select Case Cod_Base
               Case "Empresas": If ConSucursal Then ActualizarCodigoItem "Item", NumFile, "Empresas"
               Case "Accesos": ActualizarCodigo "Codigo", NumFile, Cod_Base
               Case "Acceso_Empresa": ActualizarCodigo "Codigo", NumFile, Cod_Base
               Case "Catalogo_Cuentas": If ConSucursal = False Then ActualizarCodigoItem "Codigo", NumFile, Cod_Base
               Case "Catalogo_Cursos": ActualizarCodigoItem "Curso", NumFile, Cod_Base
               Case "Catalogo_Estudiantil": ActualizarCodigoItem "CodigoE", NumFile, Cod_Base
               Case "Catalogo_Periodo_Lectivo": ActualizarTablaCompletaItem Cod_Base
               Case "Catalogo_SubCtas": ActualizarCodigoItem "Codigo", NumFile, Cod_Base
               Case "Catalogo_CxCxP": ActualizarCodigoCta "Codigo", NumFile, Cod_Base
               Case "Catalogo_Rol_Pagos": ActualizarCodigoItem "Codigo", NumFile, Cod_Base
               Case "Catalogo_Rol_Rubros": ActualizarCodigoItem "Codigo", NumFile, Cod_Base
               Case "Catalogo_Productos": ActualizarCodigoItem "Codigo_Inv", NumFile, Cod_Base
               Case "Catalogo_Materias": ActualizarCodigoItem "CodMat", NumFile, Cod_Base
               Case "Catalogo_Prestamo": ActualizarCodigoItem "CTP", NumFile, Cod_Base
               Case "Catalogo_Lineas": ActualizarCodigoItem "Codigo", NumFile, Cod_Base
               Case "Codigos": If ConSucursal = False Then ActualizarCodigoItem "Concepto", NumFile, "Codigos"
               Case "Clientes": ActualizarCodigo "Codigo", NumFile, Cod_Base
               Case "Clientes_Facturacion": ActualizarTablaCompletaItem Cod_Base
               Case "Clientes_Matriculas": ActualizarTablaCompletaItem Cod_Base
               Case "Clientes_Garantes": ActualizarMayor NumFile, Cod_Base
               Case "Clientes_Libretas": ActualizarRangoFecha NumFile, Cod_Base
               Case "Ctas_Proceso": If ConSucursal = False Then ActualizarCodigoItem "Detalle", NumFile, Cod_Base
               Case "Comprobantes": ActualizarRangoFecha NumFile, Cod_Base
               Case "Detalle_Factura": ActualizarRangoFecha NumFile, Cod_Base
               Case "Facturas": ActualizarRangoFecha NumFile, Cod_Base
               Case "Prestamos": ActualizarRangoFecha NumFile, Cod_Base
               Case "Seteos_Documentos": ActualizarTablaCompletaItem Cod_Base
               Case "Trans_Abonos": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Actas": ActualizarTablaCompletaItem Cod_Base
               Case "Trans_Air": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Asistencia": ActualizarTablaCompletaItem Cod_Base
               Case "Trans_Anulados": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Conciliacion": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Gastos_Caja": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Kardex": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Saldo_Libretas": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Libretas": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Intereses": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Llamadas": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Prestamos": ActualizarMayor NumFile, Cod_Base
               Case "Trans_Rol_Horas": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Rol_Pagos": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Rol_de_Pagos": ActualizarTablaCompletaItem Cod_Base
               Case "Trans_Bloqueos": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Fletes": ActualizarTablaCompletaItem Cod_Base
               Case "Trans_Notas": ActualizarTablaCompletaItem Cod_Base
               Case "Trans_Notas_Auxiliares": ActualizarTablaCompletaItem Cod_Base
               Case "Trans_Notas_Grado": ActualizarTablaCompletaItem Cod_Base
               Case "Trans_Entrada_Salida": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Compras": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Ventas": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Importaciones": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_Exportaciones": ActualizarRangoFecha NumFile, Cod_Base
               Case "Trans_SubCtas": ActualizarRangoFecha NumFile, Cod_Base
               Case "Transacciones": ActualizarRangoFecha NumFile, Cod_Base
             End Select
        Close #NumFile
        LstArchivo.AddItem " Proceso existoso."
        LstArchivo.Refresh
        RatonNormal
    Next IJ
  Progreso_Esperar
  Progreso_Final
  RatonNormal
  Respaldos.Caption = "MODULO DE RESPALDOS"
  NumEmpresa = AuxNumEmp
  MsgBox "Fin del Proceso"
End Sub

'No borrar hasta despuesde unos 5 años mas 2010
'''Private Sub Command6_Click()
'''  If NumEmpresa = "" Then NumEmpresa = LGrupo.Text
''''''  sSQL = "UPDATE Accesos " _
''''''       & "SET Item = '" & NumEmpresa & "' " _
''''''       & "WHERE Item <> '000' "
''''''  ConectarAdoExecute sSQL
''''  ConSubDir = False
'''  'CommonDialog1.ShowOpen   ' presentar el cuadro de diálogo común Abrir.
'''  'RutaDestino = CommonDialog1.FileName
'''  If RutaDestino <> "" Then Respaldo_Anterior
'''  Unload Respaldos
'''End Sub

Private Sub Form_Activate()
  RatonReloj
  FechaValida MBFechaI
  FechaValida MBFechaF
  LGrupo.Clear
  sSQL = "SELECT Grupo " _
       & "FROM Empresas " _
       & "WHERE Grupo <> '000' " _
       & "AND LEN(Grupo) = 3 " _
       & "GROUP BY Grupo "
  SelectData AdoQuery, sSQL
  With AdoQuery.Recordset
    Do While Not .EOF
       LGrupo.AddItem .Fields("Grupo")
       Codigo = RutaSysBases & "\DATOS\D" & .Fields("Grupo")
       Cadena = Dir(Codigo, vbDirectory)
       If Cadena = "" Then MkDir (Codigo)
      .MoveNext
    Loop
  End With
  RutaBackup = RutaSysBases
  LGrupo.Text = NumEmpresa
  GrupoEmpresa = NumEmpresa
  Codigo = Mid(MBFechaI, 1, 2) & Mid(MBFechaI, 4, 2)   ' Dia y mes de respaldo
  LstArchivo.Clear
  LstArchivo.AddItem NumEmpresa & ", Carpeta Base Anterior: " & Carpeta
  Respaldos.Caption = "MODULO DE RESPALDOS"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm Respaldos
  If CodigoUsuario = "ACCESO02" Then
'     Command6.Visible = True
     Command4.Visible = False
  Else
'     Command6.Visible = False
     Command4.Visible = True
  End If
  ConectarAdodc AdoAux
  ConectarAdodc AdoAct
  'ConectarAdodc AdoOld
  ConectarAdodc AdoQuery
End Sub

Private Sub LGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub LGrupo_LostFocus()
Dim MiArchivo As String
Dim MiRuta    As String
Dim MiNombre  As String
   GrupoEmpresa = LGrupo.Text
   LstArchivo.Clear
   MiRuta = RutaSysBases & "\DATOS\D" & GrupoEmpresa & "\"
   MiArchivo = Dir(RutaSysBases & "\DATOS\D" & GrupoEmpresa & "\*.txt")
   Do While MiArchivo <> ""
      If MiArchivo <> "." And MiArchivo <> ".." Then
         If (GetAttr(MiRuta & MiArchivo) And vbNormal) = vbNormal Then
            LstArchivo.AddItem MiArchivo
         End If
      End If
      MiArchivo = Dir
   Loop
End Sub

'''Private Sub LstTipoRespaldo_LostFocus()
'''  If LstTipoRespaldo.Selected(0) = True Then
'''     For I = 1 To 4
'''         LstTipoRespaldo.Selected(I) = False
'''     Next I
'''  End If
'''End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyF10 Then
     LstArchivo.Clear
     If File1.ListCount > 0 Then
        RatonReloj
        RutaGeneraFile = RutaSysBases & "\DATOS\R" & GrupoEmpresa & "\" & File1.List(0)
        NumFile = FreeFile
        Open RutaGeneraFile For Input As #NumFile
        Line Input #NumFile, Cadena
        Line Input #NumFile, Cadena
        AbrirArchivoSQL NumFile
        Close #NumFile
        RatonNormal
     End If
  End If
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
Dim RutaRespaldos As String
  'MsgBox "|" & CompilarString(InputBox("Cadena:", "Texto")) & "|"
  RatonReloj
  LstArchivo.Clear
  Cod_FechaI = "": Cod_FechaF = ""
  LstArchivo.Clear
  RutaRespaldos = RutaSysBases & "\DATOS\D" & GrupoEmpresa & "\"
  Cadena = Dir(RutaRespaldos & "*.txt", vbNormal)
  Do While Cadena <> ""
     If Cadena <> "." And Cadena <> ".." Then
        If (GetAttr(RutaRespaldos & Cadena) And vbNormal) = vbNormal And Val(Mid(Cadena, Len(Cadena) - 6, 3)) = 0 Then
           RutaGeneraFile = RutaRespaldos & Cadena
        End If
     End If
     Cadena = Dir
  Loop
  If RutaGeneraFile <> "" Then
  Cadena = Dir(RutaGeneraFile, vbArchive)
  If Cadena <> "" Then
    'MsgBox Cadena
     NumFile = FreeFile
     Open RutaGeneraFile For Input As #NumFile
          Line Input #NumFile, Cadena
          Line Input #NumFile, Cadena
          AbrirArchivoSQL NumFile
     Close #NumFile
  End If
  End If
  RatonNormal
  If Cod_FechaI = "" Then Cod_FechaI = FechaSistema
  If Cod_FechaF = "" Then Cod_FechaF = FechaSistema
  MBFechaI = Cod_FechaI
  MBFechaF = Cod_FechaF
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

Public Function LeerCamposTabla(Cod_FieldT1 As String, _
                                BCodigo As Variant) As Variant
Dim CodBusq1 As Variant
Dim Valor_Field As String

  Empleados = False
  CodBusq1 = "0"
  Cod_Field = Cod_FieldT1
  No_Desde = 1: No_Hasta = 1
  For I = 1 To CantCampos
   If Cod_FieldT1 <> "" Then
      Do While No_Hasta <= Len(Cod_FieldT1) And Mid(Cod_FieldT1, No_Hasta, 1) <> "|"
         No_Hasta = No_Hasta + 1
      Loop
      Valor_Field = Trim(Mid(Cod_FieldT1, No_Desde, No_Hasta - 1))
      If Valor_Field = "" Then Valor_Field = "0"
      TipoC(I).Valor = Valor_Field
      'MsgBox TipoC(I).Campo & vbCrLf & TipoC(I).Valor
      If TipoC(I).Campo = "Fecha" Then Mifecha = BuscarFecha(Format(Valor_Field, FormatoFechas))
      If TipoC(I).Campo = "Mes_No" Then NoMeses = Valor_Field
      If TipoC(I).Campo = "TP" Then TipoProc = Valor_Field
      If TipoC(I).Campo = "Credito_No" Then Contrato_No = Valor_Field
      If TipoC(I).Campo = "Item" Then NumItemTemp = Valor_Field
      If TipoC(I).Campo = "Cta" Then Cta = TipoC(I).Valor
      If TipoC(I).Campo = BCodigo Then CodBusq1 = TipoC(I).Valor
      If TipoC(I).Campo = "Periodo" Then Empleados = True
   End If
   Cod_FieldT1 = Mid(Cod_FieldT1, No_Hasta + 1, Len(Cod_FieldT1))
   No_Desde = 1: No_Hasta = 1
  Next I
  LeerCamposTabla = CodBusq1
End Function

Public Sub SetCamposTabla(FAddNew As Boolean)
Dim SetI As Integer
Dim SetJ As Integer
Dim CantCampos1 As Integer
Dim SiEncontro As Boolean

'Set Campos de tabla
  If ID_Trans <= 0 Then ID_Trans = ID_Trans + 1
  'MsgBox AdoQuery.RecordSource
  With AdoQuery.Recordset
   If FAddNew Then SetAddNew AdoQuery
   For SetJ = 0 To .Fields.Count - 1
       SiEncontro = False: SetI = 1
       Do
         If .Fields(SetJ).Name = TipoC(SetI).Campo Then
             SetFields AdoQuery, TipoC(SetI).Campo, TipoC(SetI).Valor
             SiEncontro = True
         End If
         SetI = SetI + 1
       Loop Until SetI > (UBound(TipoC) - 1)   ' CantCampos
       
       If SiEncontro = False Then
          Select Case .Fields(SetJ).Type
            Case TadBoolean
                 SetFields AdoQuery, .Fields(SetJ).Name, False
            Case TadDate, TadDate1
                 SetFields AdoQuery, .Fields(SetJ).Name, FechaSistema
            Case TadTime
                 SetFields AdoQuery, .Fields(SetJ).Name, TiempoSistema
            Case TadByte, TadInteger, TadLong, TadDouble, TadSingle, TadCurrency
                 SetFields AdoQuery, .Fields(SetJ).Name, 0
            Case TadText
                 SetFields AdoQuery, .Fields(SetJ).Name, Ninguno
            Case Else
                 SetFields AdoQuery, .Fields(SetJ).Name, Ninguno
          End Select
       End If
   Next SetJ
   SetUpdate AdoQuery
  End With
End Sub

Public Sub Porcentaje_Proceso(NombTabla As String, _
                              ContX As Long)
  Cadena = NombTabla & ": Procesando(" & Format(ContX / TotalReg, "##0%") & ") " _
          & String(ContX Mod 40, "|")
         
  ProgBarr.Mensaje_Box = NombTabla
  'Pict_Proceso PictTabla, ProgBarr
  'Respaldos.Caption = Cadena
  'MsgBox Cadena
End Sub

Public Sub EliminarRepetidosDe(EsNum As Boolean, _
                               NombreTabla As String, _
                               BCodigo As Variant)
Dim CodBusq As Variant
Dim ValorCamp As Variant
  RatonReloj
  Contador = 0
  F = 0
  NombreTabla = Trim(NombreTabla)
  BCodigo = Trim(BCodigo)
  sSQL = "SELECT * " _
       & "FROM " & NombreTabla & " "
  If EsNum Then
     sSQL = sSQL & "WHERE " & BCodigo & " <> 0 "
  Else
     sSQL = sSQL & "WHERE " & BCodigo & " <> '.' "
  End If
  sSQL = sSQL & "ORDER BY " & BCodigo & " "
  SelectAdodc AdoQuery, sSQL
  SelectAdodc AdoAct, sSQL
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          CodBusq = .Fields(BCodigo)
          Respaldos.Caption = NombreTabla & " (" & F & ") " & CodBusq & " -> " & Contador & "/" & .RecordCount
          sSQL = "SELECT * " _
               & "FROM " & NombreTabla & " "
          If EsNum Then
             sSQL = sSQL & "WHERE " & BCodigo & "=" & CodBusq & " "
          Else
             sSQL = sSQL & "WHERE " & BCodigo & "='" & CodBusq & "' "
          End If
          SelectAdodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then
             AdoAux.Recordset.MoveLast
             If AdoAux.Recordset.RecordCount > 1 Then
                sSQL = "DELETE * " _
                     & "FROM " & NombreTabla & " "
                If EsNum Then
                   sSQL = sSQL & "WHERE " & BCodigo & "=" & CodBusq & " "
                Else
                   sSQL = sSQL & "WHERE " & BCodigo & "='" & CodBusq & "' "
                End If
                ConectarAdoExecute sSQL
                F = F + 1
                SetAddNew AdoAct
                For I = 0 To AdoAct.Recordset.Fields.Count - 1
                    Codigo = AdoAct.Recordset.Fields(I).Name
                    ValorCamp = AdoAux.Recordset.Fields(I)
                    SetFields AdoAct, Codigo, ValorCamp
                Next I
                SetUpdate AdoAct
             End If
          End If
         .MoveNext
       Loop
   End If
  End With
End Sub

Public Sub Preparar_Clientes_97()
 'MsgBox AdoStrCnn1 & vbCrLf & vbCrLf & AdoStrCnn2
  Contador = 0
 'Cuentas libretas
'''  sSQL = "SELECT * " _
'''       & "FROM Cuentas " _
'''       & "WHERE Cuenta_No <> '.' " _
'''       & "ORDER BY Apellidos,Nombres "
'''  SelectAdodc AdoAct, sSQL
'''  With AdoAct.Recordset
'''   If .RecordCount > 0 Then
'''       Do While Not .EOF
'''          Contador = Contador + 1
'''          Respaldos.Caption = "Procesando " & Format(Contador / .RecordCount, "00%") & ", Codigo de Clientes en Cuentas Libretas..."
'''          SetAdoAddNew "Clientes"
'''          SetAdoFields "T", "N"
'''          SetAdoFields "Codigo", Compilar_Strg_Migracion(.Fields("Cuenta_No"))
'''          SetAdoFields "Cliente", UCase(Trim(Compilar_Strg_Migracion(.Fields("Apellidos") & " " & .Fields("Nombres"))))
'''          SetAdoFields "CI_RUC", CompilarRUC_CI(.Fields("RUC_CI"))
'''          SetAdoFields "Direccion", .Fields("Direccion")
'''          SetAdoFields "Ciudad", UCase(Compilar_Strg_Migracion(.Fields("Ciudad")))
'''          SetAdoFields "Telefono", .Fields("Telefono")
'''          SetAdoFields "Celular", .Fields("TelefonoT")
'''          SetAdoFields "FAX", .Fields("FAX")
'''          SetAdoFields "Grupo", NumEmpresa
'''          SetAdoFields "TD", "O"
'''          SetAdoUpdate
'''         .MoveNext
'''       Loop
'''   End If
'''  End With
  AdoStrCnn = AdoStrCnn2
  sSQL = "UPDATE Comprobantes SET Autorizado = '.' WHERE Autorizado <> '.' "
  ConectarAdoExecute sSQL
  Contador = 0
 'Buscar Clientes Cuenta Libretas
  sSQL = "SELECT Beneficiario,RUC_CI " _
       & "FROM Comprobantes " _
       & "WHERE Beneficiario <> '.' " _
       & "GROUP BY Beneficiario,RUC_CI " _
       & "ORDER BY Beneficiario "
  SelectAdodc AdoOld, sSQL
  With AdoOld.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          TipoBenef = "O"
          CodigoCliente = Ninguno
          Contador = Contador + 1
          Respaldos.Caption = "Actualizando Clientes(C): " & Format(Contador / .RecordCount, "00.0%")
          Codigo = .Fields("Beneficiario")
          CICliente = CompilarRUC_CI(.Fields("RUC_CI"))
          DigVerif = Digito_Verificador(CICliente)
          Select Case Tipo_RUC_CI.Tipo_Beneficiario
            Case "C": CodigoCliente = Mid(CICliente, 1, 10)
            Case "R": CodigoCliente = Chr(Val(Mid(CICliente, 1, 1)) + 65) _
                                    & Chr(Val(Mid(CICliente, 2, 1)) + 65) & Mid(CICliente, 3, 8)
            Case Else: CodigoCliente = NumEmpresa & Format(Contador, "0000000")
                       CICliente = NumEmpresa & Format(Contador, "000000")
          End Select
          sSQL = "UPDATE Comprobantes " _
               & "SET RUC_CI = '" & CICliente & "', " _
               & "Autorizado = '" & CodigoCliente & "' " _
               & "WHERE Beneficiario = '" & Codigo & "' "
          ConectarAdoExecute sSQL
         .MoveNext
       Loop
      .UpdateBatch
   End If
  End With
 'Actualizamos el RUC/CI de los Beneficiarios
  Contador = 0
  sSQL = "SELECT * " _
       & "FROM Beneficiarios " _
       & "WHERE Beneficiario <> '999999999999' " _
       & "ORDER BY Beneficiario "
  SelectAdodc AdoOld, sSQL
  With AdoOld.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          TipoBenef = "O"
          CICliente = CompilarRUC_CI(.Fields("RUC_CI"))
          DigVerif = Digito_Verificador(CICliente)
          Select Case Tipo_RUC_CI.Tipo_Beneficiario
            Case "C": CICliente = Mid(CICliente, 1, 10)
            Case "R": CICliente = Mid(CICliente, 1, 13)
            Case Else: CICliente = NumEmpresa & Format(Contador, "00000")
          End Select
          Respaldos.Caption = "Actualizando Clientes(B): " & Format(Contador / .RecordCount, "00.0%")
         .Fields("RUC_CI") = CICliente
         .MoveNext
       Loop
      .UpdateBatch
   End If
  End With
  AdoStrCnn = AdoStrCnn1
  sSQL = "DELETE * " _
       & "FROM Clientes " _
       & "WHERE Grupo = '" & NumEmpresa & "' "
  ConectarAdoExecute sSQL
 'Insertamos en la Base Actual los Beneficiarios
  AdoStrCnn = AdoStrCnn2
  sSQL = "SELECT RUC_CI,Beneficiario,Autorizado " _
       & "FROM Comprobantes " _
       & "WHERE Beneficiario <> '999999999999' " _
       & "GROUP BY RUC_CI,Beneficiario,Autorizado " _
       & "ORDER BY RUC_CI,Beneficiario,Autorizado "
  SelectAdodc AdoOld, sSQL
  AdoStrCnn = AdoStrCnn1
  Contador = 0
  With AdoOld.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          Respaldos.Caption = "Insertando Clientes(C): " & Format(Contador / .RecordCount, "00.0%")
          TipoBenef = "O"
          CICliente = CompilarRUC_CI(.Fields("RUC_CI"))
          DigVerif = Digito_Verificador(CICliente)
          SetAdoAddNew "Clientes"
          SetAdoFields "T", "N"
          SetAdoFields "Codigo", .Fields("Autorizado")
          SetAdoFields "Cliente", UCase(CompilarString(.Fields("Beneficiario")))
          SetAdoFields "CI_RUC", CICliente
          SetAdoFields "Grupo", NumEmpresa
          SetAdoFields "TD", TipoBenef
          SetAdoUpdate
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Beneficiarios " _
       & "WHERE Beneficiario <> '999999999999' " _
       & "ORDER BY Beneficiario "
  SelectAdodc AdoOld, sSQL
  Contador = 0
  With AdoOld.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          Respaldos.Caption = "Insertando Clientes(B): " & Format(Contador / .RecordCount, "00.0%")
          TipoBenef = "O"
          CICliente = CompilarRUC_CI(.Fields("RUC_CI"))
          DigVerif = Digito_Verificador(CICliente)
          SetAdoAddNew "Clientes"
          SetAdoFields "T", "N"
          SetAdoFields "Codigo", .Fields("Codigo")
          SetAdoFields "Cliente", UCase(CompilarString(.Fields("Beneficiario")))
          SetAdoFields "CI_RUC", .Fields("RUC_CI")
          SetAdoFields "Direccion", .Fields("Direccion")
          SetAdoFields "Telefono", CompilarString(.Fields("Telefono"))
          SetAdoFields "FAX", CompilarString(.Fields("FAX"))
          SetAdoFields "Celular", CompilarString(.Fields("Celular"))
          SetAdoFields "Ciudad", CompilarString(.Fields("Ciudad"))
          SetAdoFields "Grupo", NumEmpresa
          SetAdoFields "TD", TipoBenef
          SetAdoUpdate
         .MoveNext
       Loop
   End If
  End With
  Respaldos.Caption = "Actualizando los Clientes Nuevos"
  sSQL = "UPDATE Clientes SET Direccion = 'SD' WHERE Direccion = '.' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Clientes SET DirNumero = 'SN' WHERE DirNumero = '.' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Clientes SET Telefono = '022000000' WHERE Telefono = '.' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Clientes SET Celular = '099000000' WHERE Celular = '.' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Clientes SET FAX = '022000000' WHERE FAX = '.' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Clientes SET Prov = '17' WHERE Prov = '.' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Clientes SET Pais = '593' WHERE Pais = '.' "
  ConectarAdoExecute sSQL
End Sub

