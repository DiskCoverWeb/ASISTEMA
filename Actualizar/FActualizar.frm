VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form FActualizar 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATOS Y PROGRAMAS"
   ClientHeight    =   5040
   ClientLeft      =   105
   ClientTop       =   780
   ClientWidth     =   6570
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   Icon            =   "FActualizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FActualizar.frx":1297D
   ScaleHeight     =   5040
   ScaleWidth      =   6570
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImgLstFTP"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir de la actualizacion"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   "Completa"
            Description     =   ""
            Object.ToolTipText     =   "Actualiza toda la base de datos con los ejecutable"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ejecutables"
            Object.ToolTipText     =   "Actualiza solo los ejecutables"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BaseDatos"
            Object.ToolTipText     =   "Actualiza solo la base de Datos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SoloDatos"
            Object.ToolTipText     =   "Actualiza Base de Datos sin transmision"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SoloSPFN"
            Object.ToolTipText     =   "Actualiza SP y FN sin transmision"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imagenes"
            Object.ToolTipText     =   "Actualizar los Fondos y formatos"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Servidor"
            Object.ToolTipText     =   "Actualiza todas las bases del servidor actual"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Servidor_SP_FN"
            Object.ToolTipText     =   "Actualiza los SP y FN de todas las bases del servidor actual"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox TxtID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5355
         TabIndex        =   11
         Text            =   "00"
         Top             =   105
         Width           =   540
      End
   End
   Begin ComctlLib.ProgressBar ProgressBarEstado 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   1680
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame FrmBaseDatos 
      BackColor       =   &H00C0C000&
      Caption         =   "BASE DE DATOS DEL SERVIDOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1905
      Left            =   105
      TabIndex        =   9
      Top             =   3045
      Width           =   6315
      Begin VB.TextBox TxtBaseDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1485
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   210
         Width           =   6105
      End
   End
   Begin VB.Timer Timer1 
      Left            =   7875
      Top             =   630
   End
   Begin VB.ListBox LstStatud 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   105
      TabIndex        =   7
      Top             =   2100
      Width           =   6315
   End
   Begin InetCtlsObjects.Inet URLinet 
      Left            =   630
      Top             =   4620
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RemotePort      =   10222
      URL             =   "http://"
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1890
      TabIndex        =   4
      Top             =   4620
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   3885
      TabIndex        =   3
      Top             =   4515
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   5250
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSAdodcLib.Adodc AdoEmpresa 
      Height          =   330
      Left            =   2205
      Top             =   5250
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
      Caption         =   "Empresa"
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
   Begin MSAdodcLib.Adodc AdoBusqEmp 
      Height          =   330
      Left            =   105
      Top             =   5565
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
      Caption         =   "BusqEmp"
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
   Begin MSComDlg.CommonDialog Dir_Dialog 
      Left            =   105
      Top             =   4620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2205
      Top             =   5565
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
   Begin ComctlLib.ListView LstVwFTP 
      Height          =   645
      Left            =   105
      TabIndex        =   2
      Top             =   5985
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1138
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImgLstFTP"
      SmallIcons      =   "ImgLstFTP"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Archivos"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tamaño"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Modificado"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.ListBox LstTablas 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   690
      Left            =   4725
      TabIndex        =   0
      Top             =   5250
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label LblAdvertencia 
      BackColor       =   &H00800080&
      BackStyle       =   0  'Transparent
      Caption         =   "ES UNA PRUEBA DEL TEXTO"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   105
      TabIndex        =   6
      Top             =   735
      Width           =   6315
   End
   Begin ComctlLib.ImageList ImgLstFTP 
      Left            =   1260
      Top             =   4620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1C1AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1C4C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1C7E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1CAE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1CE03
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1D11D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1D40F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1DC29
            Key             =   "archivo"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1DF43
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1E25D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1E49B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FActualizar.frx":1E7B5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MActOtraBase 
         Caption         =   "Actualicar otra Base"
         Shortcut        =   ^E
      End
      Begin VB.Menu MEliminarIndce 
         Caption         =   "Eliminar Indices"
      End
      Begin VB.Menu MReindexarTablas 
         Caption         =   "Reindexar Tablas"
      End
      Begin VB.Menu MOptimitarBase 
         Caption         =   "Optimizar Base"
      End
      Begin VB.Menu MListFiles 
         Caption         =   "Lista de Archivos"
         Shortcut        =   ^U
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
Attribute VB_Name = "FActualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AdoDBMySQL As ADODB.Recordset

Dim IniIDBase As Integer
Dim FinIDBase As Integer

Dim RepresentanteEntidad As String
Dim NombreEntidad As String
Dim ftpDirUpdate As String
Dim IPDelOrdenador As String
Dim ping As cPing
Dim EsReadOnly As Boolean
'--------------------------------

Private Sub MListFiles_Click()
Dim sArchivo As String
Dim JSONFiles As String
Dim ContFile As Integer

    RatonReloj
    ContFile = 0
    JSONFiles = "{"
    sArchivo = Dir(RutaSistema & "\BASES\UPDATE_DB\*.upd")
    Do While sArchivo <> ""
       ContFile = ContFile + 1
       JSONFiles = JSONFiles & "'File_" & Format(ContFile, "000") & "' :'" & sArchivo & "'," & vbCrLf
       sArchivo = Dir
    Loop
    
    sArchivo = Dir(RutaSistema & "\BASES\UPDATE_DB\*.dbs")
    Do While sArchivo <> ""
       ContFile = ContFile + 1
       JSONFiles = JSONFiles & "'File_" & Format(ContFile, "000") & "' :'" & sArchivo & "'," & vbCrLf
       sArchivo = Dir
    Loop
    
    sArchivo = Dir(RutaSistema & "\BASES\UPDATE_DB\*.sql")
    Do While sArchivo <> ""
       ContFile = ContFile + 1
       JSONFiles = JSONFiles & "'File_" & Format(ContFile, "000") & "' :'" & sArchivo & "'," & vbCrLf
       sArchivo = Dir
    Loop
    
    JSONFiles = MidStrg(JSONFiles, 1, Len(JSONFiles) - 3) & "}"
    JSONFiles = Replace(JSONFiles, "'", """")
    JSONFiles = Replace(JSONFiles, "True", "1")
    JSONFiles = Replace(JSONFiles, "False", "0")
    JSONFiles = Replace(JSONFiles, "Verdadero", "1")
    JSONFiles = Replace(JSONFiles, "Falso", "0")
    Clipboard.Clear
    Clipboard.SetText JSONFiles
    RatonNormal
    MsgBox "Proceso Terminado, Pegue el Resultado de la copia del portapapeles"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim hInst As Long
Dim Thread As Long
            
    If Button.key <> "Salir" Then
       RatonReloj
       FActualizar.Height = Toolbar1.Top + LstStatud.Top + LstStatud.Height + 850
       LstStatud.Clear
     
       LblAdvertencia.ForeColor = &H80FFFF
       LblAdvertencia.FontBold = False
       LblAdvertencia.FontSize = 8
       LblAdvertencia.Caption = "ADVERTENCIA:" & vbCrLf & MensajeDeAdvertencia
     
       TMail.Mensaje = "Estimado(a) Cliente, le informamos que el sistema ha sido actualizado con exito." & vbCrLf
       Progreso_Barra.Mensaje_Box = Button.Description
      'Progreso_Iniciar
       Procesando = 0
       Progreso_Barra.Incremento = 0
       Progreso_Barra.Puntos = 0
       Progreso_Barra.color = 0
    End If
    
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir"                          'Salir de la actualizacion
            Progreso_Barra.Valor_Maximo = 100
            RatonNormal
            End
      Case "Completa"                       'Actualizar toda la base de datos con los ejecutable
            Progreso_Barra.Valor_Maximo = 3500
            TMail.Mensaje = TMail.Mensaje
            Actualizacion_Completa
            Enviar_Mail_Actualizacion
            Proceso_Terminado_Exitosamente
      Case "Ejecutables"                    'Actualiza solo los ejecutables
            Progreso_Barra.Valor_Maximo = 30
            Bajar_Archivos_FTP "[2]"
            TMail.Mensaje = TMail.Mensaje & "Se actualizon Los ejecutables" & vbCrLf
            Enviar_Mail_Actualizacion
            Proceso_Terminado_Exitosamente
      Case "BaseDatos"                      'Actualiza solo la base de datos
            Progreso_Barra.Valor_Maximo = 2500
            Bajar_Archivos_FTP "[1]"
            TMail.Mensaje = TMail.Mensaje & "Se actualizo solo las bases de datos" & vbCrLf
            UPD_Actualizar_SP
            Enviar_Mail_Actualizacion
            Proceso_Terminado_Exitosamente
      Case "SoloDatos"                      'Actualiza base de datos sin transmision
            Progreso_Barra.Valor_Maximo = 1900
            TMail.Mensaje = TMail.Mensaje & "Se actualizon solo los datos" & vbCrLf
            UPD_Actualizar_SP
            Enviar_Mail_Actualizacion
            Proceso_Terminado_Exitosamente
      Case "SoloSPFN"                       'Actualiza SP y FN sin transmision
            Progreso_Barra.Valor_Maximo = 150
            UPD_Actualizar_SP True
            Iniciar_Datos_Default_SP
            TMail.Mensaje = TMail.Mensaje & "Se actualizon los procedimientos por defaul" & vbCrLf
            Enviar_Mail_Actualizacion
            Proceso_Terminado_Exitosamente
      Case "Imagenes"                       'Actualiza solo la imagenes
            Progreso_Barra.Valor_Maximo = 720
            Bajar_Archivos_FTP "[3]"
            TMail.Mensaje = TMail.Mensaje & "Se actualizon los fondos y formatos del sistema" & vbCrLf
            Enviar_Mail_Actualizacion
            Proceso_Terminado_Exitosamente
      Case "Servidor"                       'Actualiza las bases de datos del servidor actual
            Progreso_Barra.Valor_Maximo = 1900
            Actualizar_Servidor
            Proceso_Terminado_Exitosamente
      Case "Servidor_SP_FN"                 'Actualiza los SP y FN de todas la bases del servidor actual
            Progreso_Barra.Valor_Maximo = 500
            Actualizar_Servidor True
            Proceso_Terminado_Exitosamente
    End Select
End Sub

Private Sub Bajar_Archivos_FTP(TypeUpdate As String)
Dim ListaDeArchivos As String
Dim sOrigen As String
Dim sDestino As String
Dim Certificados() As String
Dim Files() As String
Dim Idc As Byte
'If InStr(IP_PC.IP_PC, "192.168.") > 0 Then .servidor = "192.168.27.4" Else
On Error GoTo error_Handler
   
   FActualizar.Caption = "ESTABLECIENDO CONEXION AL SERVIDOR..."
   FActualizar.Refresh
  'If Existe_File(RutaSysBases & "\TEMP\*.zip") Then Kill RutaSysBases & "\TEMP\*.zip"
  'EsReadOnly = True
   With ftp
       Progreso_Barra.Mensaje_Box = "Conectando al servidor"
      .Mostar_Estado_FTP ProgressBarEstado, LstStatud
      .Inicializar Me
      ' MsgBox EsReadOnly
       If EsReadOnly Then
          If InStr(IPDelOrdenador, "192.168.27") Then
            .servidor = "192.168.27.3"           'Establecesmo el nombre del Servidor FTP
            .Puerto = 21
          Else
            .servidor = ftpUpSvr                 'Establecesmo el nombre del Servidor FTP
            .Puerto = ftpUpPuerto
          End If
         .Password = ftpUpPwr                    'Le establecemos la contraseña de la cuenta Ftp
         .Usuario = ftpUpUse                     'Le establecemos el nombre de usuario de la cuenta
       Else
         .servidor = ftpSvr                    'Establecesmo el nombre del Servidor FTP
         .Password = ftpPwr                    'Le establecemos la contraseña de la cuenta Ftp
         .Usuario = ftpUse                     'Le establecemos el nombre de usuario de la cuenta
         .Puerto = ftpPuerto
       End If
      'MsgBox .servidor
      'Conectamos al servidor FTP. EL label es el control donde mostrar los errores y el estado de la conexión
       If .ConectarFtp(LstStatud) = False Then
           MsgBox "No se pudo conectar"
           Exit Sub
       End If
       FActualizar.Caption = "DATOS Y PROGRAMAS: " & .servidor
       FActualizar.Refresh
      'Mostramos en el label el path del directorio actual donde estamos ubicados en el servidor
       Progreso_Barra.Mensaje_Box = .GetDirectorioActual
      'MsgBox Progreso_Barra.Mensaje_Box
      .Mostar_Estado_FTP ProgressBarEstado, LstStatud
      'Le indicamos el ListView donde se listarán los archivos
       Set .ListView = LstVwFTP
       Progreso_Barra.Mensaje_Box = "Buscando directorio en el servidor"
      .Mostar_Estado_FTP ProgressBarEstado, LstStatud

      '------------------------------------------------------------------------------------
      'Esta opcion solo actualiza la base de datos y los Store Procedure con las Functiones
      '====================================================================================
       If InStr(TypeUpdate, "[1]") Then
          Progreso_Barra.Mensaje_Box = "Eliminando Version anterior"
         .Mostar_Estado_FTP ProgressBarEstado, LstStatud
          Eliminar_Si_Existe_File RutaSistema & "\FONDOS\*.*"
          Eliminar_Si_Existe_File RutaSistema & "\FONDOS\USUARIOS\*.*"
          Eliminar_Si_Existe_File RutaSistema & "\BASES\UPDATE_DB\*.*"
          
         .CambiarDirectorio "/SISTEMA/BASES/UPDATE_DB/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
             'MsgBox LstVwFTP.ListItems(I) & vbCrLf & UCaseStrg(RightStrg(LstVwFTP.ListItems(I), 3))
              Select Case UCaseStrg(RightStrg(LstVwFTP.ListItems(I), 3))
                Case "DBS", "UPD", "TXT", "DOC", "SQL"
                     Progreso_Barra.Mensaje_Box = "Actualizando [1]: " & LstVwFTP.ListItems(I)
                    .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                    .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\BASES\UPDATE_DB\" & LstVwFTP.ListItems(I), True
                Case "ZIP"
                     Progreso_Barra.Mensaje_Box = "Actualizando: " & LstVwFTP.ListItems(I)
                    .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                    .ObtenerArchivo LstVwFTP.ListItems(I), RutaSysBases & "\TEMP\" & LstVwFTP.ListItems(I), True
              End Select
          Next I
         'Insertamos fondos nuevos de raiz
         .CambiarDirectorio "/SISTEMA/FONDOS/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              Select Case UCaseStrg(RightStrg(LstVwFTP.ListItems(I), 3))
                Case "GIF", "JPG", "PNG"
                     Progreso_Barra.Mensaje_Box = "Actualizando: FONDOS\"
                    .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                    .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\FONDOS\" & LstVwFTP.ListItems(I), True
              End Select
          Next I
       End If
       
      '------------------------------------------
      'Esta opcion solo actualiza los ejecutables
      '==========================================
       If InStr(TypeUpdate, "[2]") Then
         'Borramos archivos antiguos
          Eliminar_Si_Existe_File RutaSistema & "\JAVASCRIPT\*.*"
          
         .CambiarDirectorio "/SISTEMA"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              Select Case UCaseStrg(RightStrg(LstVwFTP.ListItems(I), 3))
                Case "EXE", "JPG", "PNG", "GIF"
                     Progreso_Barra.Mensaje_Box = "Eliminando: " & LstVwFTP.ListItems(I)
                    .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                     Eliminar_Si_Existe_File RutaSistema & "\" & LstVwFTP.ListItems(I)
              End Select
              Cadena = Cadena & LstVwFTP.ListItems(I) & vbCrLf
          Next I
          
         .CambiarDirectorio "/SISTEMA/JAVASCRIPT/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              Progreso_Barra.Mensaje_Box = "Eliminando: " & LstVwFTP.ListItems(I)
             .Mostar_Estado_FTP ProgressBarEstado, LstStatud
              Eliminar_Si_Existe_File RutaSistema & "\" & LstVwFTP.ListItems(I)
              Cadena = Cadena & LstVwFTP.ListItems(I) & vbCrLf
          Next I
          
         'Copiamos nuevos archivos del servidor
         .CambiarDirectorio "/SISTEMA"
         .ListarArchivos
          Cadena = ""
          For I = 1 To LstVwFTP.ListItems.Count
              Select Case UCaseStrg(RightStrg(LstVwFTP.ListItems(I), 3))
                Case "PNG", "JPG", "GIF"
                 Cadena = Cadena & "Copying: " & LstVwFTP.ListItems(I) & vbCrLf
                 Progreso_Barra.Mensaje_Box = "Actualizando: " & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\" & LstVwFTP.ListItems(I), True
              End Select
          Next I
          'Sleep 5000
          
         .CambiarDirectorio "/SISTEMA/JAVASCRIPT/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              Progreso_Barra.Mensaje_Box = "Actualizando: " & LstVwFTP.ListItems(I)
             .Mostar_Estado_FTP ProgressBarEstado, LstStatud
             .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\JAVASCRIPT\" & LstVwFTP.ListItems(I), True
          Next I
          
         .CambiarDirectorio "/SISTEMA/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              If UCaseStrg(RightStrg(LstVwFTP.ListItems(I), 3)) = "EXE" Then
                Cadena = Cadena & "Copying: " & LstVwFTP.ListItems(I) & vbCrLf
                 Progreso_Barra.Mensaje_Box = "Actualizando: " & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\" & LstVwFTP.ListItems(I), True
              End If
          Next I
         'Conectamos la nueva Base de Datos para sacar los Certificados del servidor
          Idc = 0
          sSQL = "SELECT Empresa, Ruta_Certificado " _
               & "FROM Empresas " _
               & "WHERE Ruta_Certificado LIKE '%P12' " _
               & "ORDER BY Empresa "
          Select_Adodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then
             Do While Not AdoAux.Recordset.EOF
                RutaDocumentos = RutaSistema & "\CERTIFIC\" & AdoAux.Recordset.Fields("Ruta_Certificado")
                If Len(Dir$(RutaDocumentos)) = 0 Then
                   ReDim Preserve Certificados(Idc) As String
                   Certificados(Idc) = AdoAux.Recordset.Fields("Ruta_Certificado")
                   Idc = Idc + 1
                End If
                AdoAux.Recordset.MoveNext
             Loop
          End If

          If Idc > 0 Then
             RatonReloj
            .CambiarDirectorio "/SISTEMA/CERTIFIC/"
            .ListarArchivos
             For I = 1 To LstVwFTP.ListItems.Count
                 For J = 0 To UBound(Certificados)
                     If Certificados(J) = LstVwFTP.ListItems(I) Then
                        Progreso_Barra.Mensaje_Box = "Actualizando: " & LstVwFTP.ListItems(I)
                       .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                       .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\CERTIFIC\" & LstVwFTP.ListItems(I), True
                     End If
                 Next J
             Next I
             RatonNormal
          End If
       End If

      '=======================================================
      'Esta opcion solo actualiza los fondos, formatos y logos
      '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       If InStr(TypeUpdate, "[3]") Then
         'Borrammos fondos antiguos de cada mes
          For J = 1 To 12
             .CambiarDirectorio "/SISTEMA/FONDOS/M" & Format(J, "00") & "/"
             .ListarArchivos
              For I = 1 To LstVwFTP.ListItems.Count
                  If Len(LstVwFTP.ListItems(I)) > 3 Then
                     Progreso_Barra.Mensaje_Box = "Eliminando: FONDOS\M" & Format(J, "00") & "\" & LstVwFTP.ListItems(I)
                    .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                     Eliminar_Si_Existe_File RutaSistema & "\FONDOS\M" & Format(J, "00") & "\" & LstVwFTP.ListItems(I)
                  End If
              Next I
          Next J
         
         .CambiarDirectorio "/SISTEMA/LOGOS/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              If Len(LstVwFTP.ListItems(I)) > 3 Then
                 Progreso_Barra.Mensaje_Box = "Eliminando: LOGOS\" & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                 Eliminar_Si_Existe_File RutaSistema & "\LOGOS\" & LstVwFTP.ListItems(I)
              End If
          Next I
          
         .CambiarDirectorio "/SISTEMA/FOTOS/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              If Len(LstVwFTP.ListItems(I)) > 3 Then
                 Progreso_Barra.Mensaje_Box = "Eliminando: FOTOS\" & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                 Eliminar_Si_Existe_File RutaSistema & "\FOTOS\" & LstVwFTP.ListItems(I)
              End If
          Next I
          
         .CambiarDirectorio "/SISTEMA/FORMATOS/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              If Len(LstVwFTP.ListItems(I)) > 3 Then
                 Progreso_Barra.Mensaje_Box = "Eliminando: FORMATOS\" & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                 Eliminar_Si_Existe_File RutaSistema & "\FORMATOS\" & LstVwFTP.ListItems(I)
              End If
          Next I
                    
         'Insertamos los fondos por Mes
          For J = 1 To 12
             .CambiarDirectorio "/SISTEMA/FONDOS/M" & Format(J, "00") & "/"
             .ListarArchivos
             'Cadena = "/SISTEMA/FONDOS/M" & Format(J, "00") & "/" & vbCrLf
              For I = 1 To LstVwFTP.ListItems.Count
                 'Cadena = Cadena & LstVwFTP.ListItems(I) & vbCrLf
                  If Len(LstVwFTP.ListItems(I)) > 3 Then
                     Progreso_Barra.Mensaje_Box = "Actualizando: FONDOS\M" & Format(J, "00") & "\" & LstVwFTP.ListItems(I)
                    .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                    .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\FONDOS\M" & Format(J, "00") & "\" & LstVwFTP.ListItems(I), True
                  End If
              Next I
             'MsgBox Cadena
          Next J
          
         .CambiarDirectorio "/SISTEMA/LOGOS/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              If Len(LstVwFTP.ListItems(I)) > 3 Then
                 Progreso_Barra.Mensaje_Box = "Actualizando: LOGOS\" & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\LOGOS\" & LstVwFTP.ListItems(I), True
              End If
          Next I
          
         .CambiarDirectorio "/SISTEMA/FOTOS/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              If Len(LstVwFTP.ListItems(I)) > 3 Then
                 Progreso_Barra.Mensaje_Box = "Actualizando: FOTOS\" & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\FOTOS\" & LstVwFTP.ListItems(I), True
              End If
          Next I
          
         .CambiarDirectorio "/SISTEMA/FORMATOS/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              If Len(LstVwFTP.ListItems(I)) > 3 Then
                 Progreso_Barra.Mensaje_Box = "Actualizando: FORMATOS\" & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\FORMATOS\" & LstVwFTP.ListItems(I), True
              End If
          Next I
       End If
       
      '_________________________________________________________
      
      'Esta opcion solo actualiza los Store Procedure y Function
      '_________________________________________________________
       If InStr(TypeUpdate, "[4]") Then
          Progreso_Barra.Mensaje_Box = "Eliminando Version anterior"
         .Mostar_Estado_FTP ProgressBarEstado, LstStatud
          Eliminar_Si_Existe_File RutaSistema & "\BASES\UPDATE_DB\*.sql"

         .CambiarDirectorio "/SISTEMA/BASES/UPDATE_DB/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
             'MsgBox LstVwFTP.ListItems(I) & vbCrLf & UCaseStrg(RightStrg(LstVwFTP.ListItems(I), 3))
              If UCaseStrg(RightStrg(LstVwFTP.ListItems(I), 3)) = "SQL" Then
                 Progreso_Barra.Mensaje_Box = "Actualizando: " & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\BASES\UPDATE_DB\" & LstVwFTP.ListItems(I), True
              End If
          Next I
       End If
      .Desconectar
   End With
   
''''  '============================================================
''''  '= Subir al Servidor ERP si es la actualizacion a las nubes =
''''  '============================================================
''''   If strIPServidor = "db.diskcoversystem.com" Then
''''     'Lista de archivos a subir
''''      'MsgBox "Subir al DB"
''''      Contador = 0
''''      ListaDeArchivos = Dir(RutaSistema & "\BASES\UPDATE_DB\*.*", vbNormal)
''''      Do While ListaDeArchivos <> ""
''''         If ListaDeArchivos <> "." And ListaDeArchivos <> ".." Then
''''            ReDim Preserve Files(Contador) As String
''''            Files(Contador) = ListaDeArchivos
''''            Contador = Contador + 1
''''         End If
''''         ListaDeArchivos = Dir
''''      Loop
''''
''''     'Empieza a subir los archivos al servidor ERP
''''      With ftp
''''           Progreso_Barra.Mensaje_Box = "Conectando al servidor"
''''          .Mostar_Estado_FTP ProgressBarEstado, LstStatud
''''          .Inicializar Me
''''          .Password = ftpPwrLinode                    'Le establecemos la contraseña de la cuenta Ftp
''''          .Usuario = ftpUseLinode                     'Le establecemos el nombre de usuario de la cuenta
''''          .servidor = ftpSvrLinode                    'Establecesmo el nombre del Servidor FTP
''''          .Puerto = 21
''''          'MsgBox .servidor
''''          'Conectamos al servidor FTP. EL label es el control donde mostrar los errores y el estado de la conexión
''''           If .ConectarFtp(LstStatud) = False Then
''''               MsgBox "Error (" & Err.Number & ") " & Err.Description & vbCrLf & "No se pudo conectar"
''''               Exit Sub
''''           End If
''''           FActualizar.Caption = "DATOS Y PROGRAMAS: " & .servidor
''''           FActualizar.Refresh
''''          'Mostramos en el label el path del directorio actual donde estamos ubicados en el servidor
''''           Progreso_Barra.Mensaje_Box = .GetDirectorioActual
''''          'MsgBox Progreso_Barra.Mensaje_Box
''''          .Mostar_Estado_FTP ProgressBarEstado, LstStatud
''''          'Le indicamos el ListView donde se listarán los archivos
''''           Set .ListView = LstVwFTP
''''           Progreso_Barra.Mensaje_Box = "Buscando directorio en el servidor"
''''          .Mostar_Estado_FTP ProgressBarEstado, LstStatud
''''
''''          '------------------------------------------------------------------------------------
''''          'Realizamos la subida del archivo
''''          '====================================================================================
''''           If InStr(TypeUpdate, "[1]") Or InStr(TypeUpdate, "[4]") Then
''''             'Eliminar_Si_Existe_File RutaSistema & "\FONDOS\*.jpg"
''''             '.EliminarArchivo
''''              For I = 0 To UBound(Files)
''''                  sOrigen = RutaSistema & "\BASES\UPDATE_DB\" & Files(I)
''''                  sDestino = "/files/UPDATE_DB/" & Files(I)
''''                  'MsgBox sOrigen & vbCrLf & sDestino
''''                  Progreso_Barra.Mensaje_Box = "Subiendo a DB: " & Files(I)
''''                 .Mostar_Estado_FTP ProgressBarEstado, LstStatud
''''                 .SubirArchivo sOrigen, sDestino, True
''''              Next I
''''           End If
''''          .Desconectar
''''      End With
''''   End If
   RatonNormal
Exit Sub
error_Handler:
     MsgBox Err.Description, vbCritical
     RatonNormal
End Sub

Public Sub Actualizacion_Completa()
Dim Si_Actualiza As Boolean
Dim Nombre_Key As String
 
 FActualizar.Caption = Modulo & ": " & strIPServidor & " - " & strNombreBaseDatos
 Si_Actualiza = True
 If Si_Actualiza Then
   'Empezamos a bajar la actualizacion del servidor en las nubes
    Bajar_Archivos_FTP "[1][2]"
    RatonReloj
    Progreso_Barra.Mensaje_Box = "PROGRESO DEL RESPALDO"
    UPD_Actualizar_SP
  Else
    Mensajes = "LO SIENTO NO PODER ACTUALIZAR EL SISTEMA, USTED NO ESTA LEGALIZADO " _
             & "O YA SE VENCIO SU CONTRATO, LLAME AL 593-09-8910-5300" & vbCrLf & vbCrLf _
             & "O ENVIE UN MAILS A: diskcoversystem@msn.com PARA SU LEGALIZACION" & vbCrLf & vbCrLf & vbCrLf _
             & "DESEA LEGALIZAR SU CLAVE"
    Titulo = "LEGALIZACION DE CLAVE"
    If BoxMensaje = vbYes Then
       ''fLeer_Campos_Key LineasLogIn(4)
       Nombre_Key = ""
       For I = 1 To CantCampos
           If (10 <= Len(TipoC(I).Valor) And Len(TipoC(I).Valor) <= 13) And IsNumeric(TipoC(I).Valor) Then
              Nombre_Key = TipoC(I).Valor
              I = CantCampos + 1
           End If
       Next I
       RutaOrigen = RutaSistema & "\FORMATOS\LOGINSYSTEM.KEY"
       Dir_Dialog.Filename = RutaSysBases & "\Key_" & Nombre_Key & ".Zip"
       Dir_Dialog.InitDir = RutaSysBases & "\"
       RutaDestino = UCase(SelectDialogFile())
       '''NombreArchivoZip = RutaDestino
       '''Empaquetar_Archivos_Zip
    End If
  End If
End Sub

Private Sub Form_Activate()
   UPD_Listar_Tablas LstTablas
   TMail.ListaMail = 0
   IniIDBase = 0
   RatonNormal
'  CommandButton2.SetFocus
End Sub

Private Sub Form_Initialize()
   SetErrorMode 2
   InitCommonControls
End Sub

Private Sub Form_Load()
Dim nFrames As Long
Dim NumFile As Integer
Dim NumPos As Long
Dim RutaGeneraFile As String
Dim LineaTexto As String
Dim MiArchivo, MiRuta, MiNombre
Dim AnchoTemp As Single
Dim HayCnn As Boolean

    RatonReloj
    CentrarForm FActualizar
   '------------------
    Set ping = New cPing
    EsReadOnly = True
    IPDelOrdenador = ping.IP_Del_PC()
   '-----------------------------------
    FActualizar.Height = Toolbar1.Top + LstStatud.Top + LstStatud.Height + 820
    
    MDI_X_Max = Screen.width - 150
    MDI_Y_Max = Screen.Height - 1850
    
    
   'Redondear_Cuadro FActualizar, 25
    ConSubDir = False
    RutaDestino = UCase$(Left$(CurDir$, 2))
    RutaSubDirTemp = RutaDestino
    RutaUpdate = RutaDestino & "\SISTEMA"
    RutaSistema = RutaDestino & "\SISTEMA"
    RutaSysBases = RutaDestino & "\SYSBASES"
    RutaEmpresa = UCase(RutaSistema & "\EMPRESA")
    RutaEmpresaOld = UCase(RutaSistema & "\EMPRESA")
    ChDir RutaSistema
    IngresarClave = True
    
   'MsgBox RutaSistema & "\JAVASCRIPT"
   '-----------------------------------------------------------------------------------------
   'Determinamos si existen carpetas nuevas del sistema
    If Not Existe_Carpeta(RutaSistema & "\JAVASCRIPT") Then MkDir RutaSistema & "\JAVASCRIPT"
   '-----------------------------------------------------------------------------------------
   'MODULOS
    NumModulo = "98"
    Modulo = "UPDATE"
    MenuDeModulos = True
    TiempoSistema = Time
    Timer1.Interval = 10000
   
   'Determinar que tipo de bases utilizamos
    Evaluar = False
    SQL_Server = True
    Conectar_Base_Datos
  
   'Verificamos si la base esta en Microsoft Access o en SQL Server 7.0
    FechaSistema = Format(Date, FormatoFechas)
    NombrePais = "Ecuador"
    NombreCiudad = "Quito"
    RUC = "9999999999999"
    NumEmpresa = "000"
    IDEUsuario = "ACCESO99"
    CodigoUsuario = "ACCESO99"
    NombreUsuario = "Actualizacion del Sistema"
    Empresa = "ACTUALIZACION DE BASES"
    RazonSocial = "DISKCOVER SYSTEM"
    NombreComercial = "ACTUALIZACION DE BASES"
    NombreContador = "Actualizacion del Sistema"
    RUC_Contador = RUC
    NombreGerente = "WALTER VACA PRIETO"
    EmailProcesos = CorreoUpdate
    EmailEmpresa = CorreoUpdate
    Telefono1 = "09-9965-4196"
    Telefono2 = "09-8910-5300"
    Periodo_Contable = Ninguno
    LogoTipo = UCase(RutaSistema & "\LOGOS\DEFAULT.GIF")
    NLogoTipo = "DiskCover"
    Direccion = "www.diskcoversystem.com"
    Mifecha = FechaSistema
    
    
    'HayCnn = Get_WAN_IP
    IP_PC.InterNet = Get_Internet
    
   ' Acceso_IP_PCs_SP_MySQL Si_No
   '|--=:******* CONECCON A MYSQL *******:=--|
     Datos_Iniciales_Entidad_SP_MySQL
   '|--=:******* --------.------- *******:=--|
 
    TMail.de = CorreoDiskCover
    TMail.ListaMail = 0
    LblAdvertencia.Caption = "ADVERTENCIA:" & vbCrLf & MensajeDeAdvertencia
    LstStatud.Clear
    Set ftp = New cFTP
    Dir1.Path = RutaEmpresa
    File1.Filename = Dir1.Path & "\*.MDB"
    
    ConectarAdodc AdoAux
    ConectarAdodc AdoBusqEmp
    ConectarAdodc AdoEmpresa
      
    sSQL = "SELECT Aplicacion " _
         & "FROM Modulos " _
         & "WHERE Modulo = 'VS' "
    Select_Adodc AdoAux, sSQL
    If AdoAux.Recordset.RecordCount > 0 Then
       Version_Sistema = "VERSION ACTUAL " & AdoAux.Recordset.Fields("Aplicacion") & " - Ver. 5.20 "
    Else
       Version_Sistema = "Teléfono del proveedor: " & vbCrLf & "(+593) 09-9965-4196/09-8910-5300. "
    End If
    FActualizar.Caption = Modulo & ": " & strIPServidor & " - " & strNombreBaseDatos
    LstStatud.Clear
    LstStatud.AddItem "D I S K C O V E R   S Y S T E M"
    LstStatud.AddItem Version_Sistema
    LstStatud.AddItem "PRISMANET PROFESIONAL S.A."
    RatonNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set ftp = Nothing
   Control_Procesos "Q", "Salir Modulo de Update"
   End
End Sub

Private Sub MActOtraBase_Click()
  FOtraBase.Top = FActualizar.Top + Toolbar1.Height + 150
  FOtraBase.Left = FActualizar.Left + 150
  FOtraBase.Show 1
End Sub

Private Sub MEliminarIndce_Click()
  Eliminar_Indices_SP
  MsgBox "Proceso Terminado"
End Sub

Private Sub MOptimitarBase_Click()
   Optimizar_Memoria
End Sub

Private Sub MReindexarTablas_Click()
  Ejecutar_SP "sp_Eliminar_Indices_Temporales", ""
  Ejecutar_SP "sp_Crear_Indices", ""
  MsgBox "Proceso Terminado"
End Sub

Private Sub MSalir_Click()
   End
End Sub

'''Public Function ConverType(CampoType As Integer) As String
'''Dim TipoCampo As String
'''  If SQL_Server Then
'''     TipoCampo = SinEspaciosIzq(TablaNew(CampoType).TipoSQL)
'''  Else
'''     TipoCampo = SinEspaciosIzq(TablaNew(CampoType).TipoAccess)
'''  End If
'''  Select Case TipoCampo
'''    Case "BIT", "TINYINT", "SMALLINT", "INT", "BYTE", "SHORT", "LONG", "INTEGER"
'''         ConverType = "CInt(" & TablaNew(CampoType).Campo & ")"
'''    Case "REAL", "FLOAT", "MONEY", "DECIMAL", "SINGLE", "DOUBLE", "CURRENCY"
'''         ConverType = "CSng(" & TablaNew(CampoType).Campo & ")"
'''    Case "NTEXT", "NVARCHAR", "TEXT", "LONGTEXT"
'''         ConverType = "CStr(" & TablaNew(CampoType).Campo & ")"
'''    Case Else
'''         ConverType = TablaNew(CampoType).Campo
'''  End Select
'''End Function

'''Public Sub Leer_Campos_Key(LineaCampos As String)
'''    Cod_Emp = "": Cod_Base = "": Cod_Field = ""
'''    CantCampos = 0
'''    For I = 1 To Len(LineaCampos)
'''        If Mid$(LineaCampos, I, 1) = "|" Then CantCampos = CantCampos + 1
'''    Next I
'''    ReDim TipoC(CantCampos) As Campos_Tabla
'''    No_Desde = 1: No_Hasta = 1
'''    Cadena = LineaCampos
'''    For I = 1 To CantCampos
'''        Do
'''           No_Hasta = No_Hasta + 1
'''        Loop Until Mid$(Cadena, No_Hasta, 1) = "|"
'''        TipoC(I).Valor = Trim$(Mid$(Cadena, No_Desde, No_Hasta - 1))
'''        Cadena = Mid$(Cadena, No_Hasta + 1, Len(Cadena))
'''        No_Desde = 1: No_Hasta = 1
'''    Next I
'''End Sub

Public Sub UPD_Actualizar_SP(Optional Solo_FN_SP As Boolean)
Dim AdoDBTXT As ADODB.Recordset

Dim IdFile As Long
Dim CadenaTime As String
Dim Extension As String
Dim TextoFile As String
Dim FileOrigen As String
Dim FileDestino As String
Dim Directorio As String
Dim Files() As String

On Error GoTo error_Handler

   'Actualizando archivo de la nueva version del sistemas en las Bases de Datos y los SP, FN
    MiTiempo = Time
    ProgressBarEstado.Max = 100
    ProgressBarEstado.value = 0
    CadenaTime = ""
    LstStatud.AddItem "Estableciendo conexion al servidor de datos"
    LstStatud.Refresh
    
   'Determinar cuales son las tablas fijas que se van a actualizar en un Array
    UpLoad_Update_Server Solo_FN_SP
   
   'MsgBox Solo_FN_SP
   'Procedemos a ejecutar el SP que actualizara SP, FN y las tablas
    LstStatud.AddItem "Procesando actualizacion en: " & strNombreBaseDatos
    LstStatud.Text = "Procesando actualizacion en: " & strNombreBaseDatos
    LstStatud.Refresh
    
   'Store Procedure que procesa la version nueva de la actualizacion
    If Solo_FN_SP Then
       Actualizar_SP_FN_SP
    Else
      'Empezamos actualizar
       Actualizar_Base_Datos_SP
       
      'Codigos Catalogo Ctas_Proceso
       LstStatud.AddItem "Determinando Duplicados de: Ctas_Proceso"
       LstStatud.Refresh
       Eliminar_Duplicados_SP "Ctas_Proceso", "Periodo, Item, Detalle"

      'Codigos Catalogo Seteos_Documentos
       LstStatud.AddItem "Determinando Duplicados de: Seteos Documentos"
       LstStatud.Refresh
       Eliminar_Duplicados_SP "Seteos_Documentos", "Item, TP, Campo"

      'Eliminar Duplicados en el Catalogo de Cuentas
       LstStatud.AddItem "Determinando Duplicados de: Catalogo de Cuentas"
       LstStatud.Refresh
       Eliminar_Duplicados_SP "Catalogo_Cuentas", "Codigo"
    End If
   '---------------------------------------------------------------------------------
    If IP_PC.InterNet Then
       RatonReloj
       sSQL = "SELECT * " _
            & "FROM lista_estados " _
            & "WHERE Estado <> '.' " _
            & "ORDER BY ID,Estado "
       Select_AdoDB_MySQL AdoRegMySQL, sSQL
       With AdoRegMySQL
        If .RecordCount > 0 Then
            Do While Not .EOF
               sSQL = "SELECT Tipo_Referencia " _
                    & "FROM Tabla_Referenciales_SRI " _
                    & "WHERE Tipo_Referencia = 'ESTADO EMPRESA' " _
                    & "AND Codigo = '" & .Fields("Estado") & "' "
               Select_AdoDB AdoReg, sSQL
               If AdoReg.RecordCount <= 0 Then
                  SQL1 = "INSERT INTO Tabla_Referenciales_SRI (Tipo_Referencia, Codigo, Descripcion) " _
                       & "VALUES ('ESTADO EMPRESA', '" & .Fields("Estado") & "', '" & .Fields("Descripcion") & "');"
                  Ejecutar_SQL_SP SQL1
               End If
               AdoReg.Close
              .MoveNext
            Loop
        End If
       End With
       AdoRegMySQL.Close
    End If
  
    ProgressBarEstado.value = ProgressBarEstado.Max
    CadenaTime = CadenaTime & Format(Time - MiTiempo, FormatoTimes) & vbCrLf
    LstStatud.AddItem "FIN DEL PROCESO DE ACTUALIZACION [" & CadenaTime & "]"
    LstStatud.Refresh
    RatonNormal
Exit Sub
error_Handler:
    MsgBox Err.Description, vbCritical
    RatonNormal
End Sub

Public Sub UpLoad_Update_Server(Optional Solo_FN_SP As Boolean)
Dim AdoDBTXT As ADODB.Recordset
Dim IdFile As Long
Dim CadenaTime As String
Dim Extension As String
Dim TextoFile As String
Dim FileOrigen As String
Dim FileDestino As String
Dim Directorio As String
Dim Files() As String
On Error GoTo error_Handler

   'Determinar cuales son las tablas fijas que se van a actualizar en un Array
    Contador = 0
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    LstStatud.AddItem "Determinando la informacion para la actualuzacion"
    LstStatud.Refresh
    
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB "DROP PROCEDURE IF EXISTS sp_Ejecutar_SQL;"

    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Ejecutar_SQL.StoredProcedure.sql"), True
    
    If Not Solo_FN_SP Then
      'upd: Archivos de Actualizacion de Tablas nuevas o campos de actualizar
       Directorio = Dir(RutaSistema & "\BASES\UPDATE_DB\*.upd", vbNormal)
       Do While Directorio <> ""
          If Directorio <> "." And Directorio <> ".." Then
              If Directorio <> "Actualizacion.upd" Then
                 LstStatud.AddItem strNombreBaseDatos & ": Archivo '" & Directorio & "' a subir"
                 LstStatud.Text = strNombreBaseDatos & ": Archivo '" & Directorio & "' a subir"
                 LstStatud.Refresh
                 ReDim Preserve Files(Contador) As String
                 Files(Contador) = Directorio
                 Contador = Contador + 1
              End If
          End If
          Directorio = Dir
       Loop
      'dbs: Archivos de bases de datos por default en la nueva actualizacion
       Directorio = Dir(RutaSistema & "\BASES\UPDATE_DB\*.dbs", vbNormal)
       Do While Directorio <> ""
          If Directorio <> "." And Directorio <> ".." Then
             If Directorio <> "Actualizacion.dbs" Then
                LstStatud.AddItem strNombreBaseDatos & ": Archivo '" & Directorio & "' a subir"
                LstStatud.Text = strNombreBaseDatos & ": Archivo '" & Directorio & "' a subir"
                LstStatud.Refresh
                ReDim Preserve Files(Contador) As String
                Files(Contador) = Directorio
                Contador = Contador + 1
             End If
          End If
          Directorio = Dir
       Loop
    End If
   'sql: Archivos de Funciones y Procedimientos de la nueva actualizacion
    Directorio = Dir(RutaSistema & "\BASES\UPDATE_DB\*.sql", vbNormal)
    Do While Directorio <> ""
       If Directorio <> "." And Directorio <> ".." Then
          LstStatud.AddItem strNombreBaseDatos & ": Archivo '" & Directorio & "' a subir"
          LstStatud.Text = strNombreBaseDatos & ": Archivo '" & Directorio & "' a subir"
          LstStatud.Refresh
          ReDim Preserve Files(Contador) As String
          Files(Contador) = Directorio
          Contador = Contador + 1
       End If
       Directorio = Dir
    Loop
    ProgressBarEstado.Max = ProgressBarEstado.Max + Contador + 20
    
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_SP "DROP TABLE IF EXISTS Actualizacion;"

    sSQL = "CREATE TABLE Actualizacion(" _
         & "Archivo VARCHAR(100) NULL, " _
         & "Extension VARCHAR(3) NULL, " _
         & "Documento NVARCHAR(MAX) NULL, " _
         & "Con_Item BIT NULL, " _
         & "Con_Periodo BIT NULL, " _
         & "ID INT IDENTITY NOT NULL PRIMARY KEY);"
    Ejecutar_SQL_SP sSQL
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB "DROP PROCEDURE IF EXISTS sp_Actualizar_Base_Datos;"
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB "DROP PROCEDURE IF EXISTS sp_Ejecutar_SQL;"
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB "DROP PROCEDURE IF EXISTS sp_Eliminar_Indices_Temporales;"
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB "DROP PROCEDURE IF EXISTS sp_Leer_Archivo_Plano;"
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB "DROP PROCEDURE IF EXISTS sp_Actualizar_SP_FN;"

    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Ejecutar_SQL.StoredProcedure.sql"), True
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Eliminar_Indices_Temporales.StoredProcedure.sql"), True
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Leer_Archivo_Plano.StoredProcedure.sql"), True
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Actualizar_Base_Datos.StoredProcedure.sql"), True
    ProgressBarEstado.value = ProgressBarEstado.value + 1
    Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Actualizar_SP_FN.StoredProcedure.sql"), True
   'Empezamos a subir la informacion de los archivos a la TABLA de Actualizacion, para luego ejecutar el SP que empieza la actualizacion
    LstStatud.AddItem "Subiendo informacion al servidor"
    LstStatud.Refresh
   'MsgBox UBound(Files)
    For I = 0 To UBound(Files)
        TextoFile = ""
        FileOrigen = Files(I)
        Extension = RightStrg(FileOrigen, 3)
        ProgressBarEstado.value = ProgressBarEstado.value + 1
        LstStatud.AddItem strNombreBaseDatos & ": Subiendo datos de " & FileOrigen
        LstStatud.Text = strNombreBaseDatos & ": Subiendo datos de " & FileOrigen
        LstStatud.Refresh
        FileDestino = RutaSistema & "\BASES\UPDATE_DB\" & FileOrigen
        TextoFile = Leer_Archivo_Plano(FileDestino)
        TextoFile = Replace(TextoFile, vbCr, "[CR]")
        TextoFile = Replace(TextoFile, vbLf, "[LF]")
        TextoFile = Replace(TextoFile, "'", "[`]")
        TextoFile = Replace(TextoFile, "#", "[N]")
        TextoFile = Replace(TextoFile, """", "[DC]")
        If Extension = "sql" Then
           IdFile = InStr(TextoFile, "CREATE")
           If IdFile > 0 Then
              TextoFile = MidStrg(TextoFile, IdFile, Len(TextoFile))
              For K = Len(TextoFile) To Len(TextoFile) - 10 Step -1
                  If MidStrg(TextoFile, K, 2) = "GO" Then J = K
              Next K
              TextoFile = MidStrg(TextoFile, 1, J - 1)
           End If
        End If
        FileOrigen = MidStrg(FileOrigen, 1, Len(FileOrigen) - 4)
        
        Ejecutar_SQL_SP "INSERT INTO Actualizacion (Archivo, Extension, Documento) VALUES ('" & FileOrigen & "', '" & Extension & "', '" & TextoFile & "')"
    Next I
    
    Ejecutar_SQL_SP ("UPDATE Actualizacion SET Documento = REPLACE(Documento,'[CR]',CHAR(13))")
    Ejecutar_SQL_SP ("UPDATE Actualizacion SET Documento = REPLACE(Documento,'[LF]',CHAR(10))")
    Ejecutar_SQL_SP ("UPDATE Actualizacion SET Documento = REPLACE(Documento,'[DC]',CHAR(34))")
    Ejecutar_SQL_SP ("UPDATE Actualizacion SET Documento = REPLACE(Documento,'[`]',CHAR(39))")
    Ejecutar_SQL_SP ("UPDATE Actualizacion SET Documento = REPLACE(Documento,'[N]',CHAR(35))")
    
'''    sSQL = "SELECT Archivo, Extension, Documento " _
'''         & "FROM Actualizacion " _
'''         & "WHERE LEN(Documento) > 1 " _
'''         & "ORDER BY Archivo "
'''    Select_AdoDB AdoDBTXT, sSQL
'''    With AdoDBTXT
'''     If .RecordCount > 0 Then
'''         Do While Not .EOF
'''            'ProgressBarEstado.value = ProgressBarEstado.value + 1
'''            LstStatud.Text = "Actualizando signos especiales: " & .fields("Archivo")
'''            LstStatud.Refresh
'''            TextoFile = .fields("Documento")
'''            TextoFile = Replace(TextoFile, "[CR]", vbCr)
'''            TextoFile = Replace(TextoFile, "[LF]", vbLf)
'''            TextoFile = Replace(TextoFile, "[`]", "'")
'''            TextoFile = Replace(TextoFile, "[N]", "#")
'''            TextoFile = Replace(TextoFile, "[DC]", """")
'''           .fields("Documento") = TextoFile
'''           .MoveNext
'''         Loop
'''        .UpdateBatch
'''     End If
'''    End With
'''    AdoDBTXT.Close

   'Procedemos a ejecutar el SP que actualizara SP, FN y las tablas
    LstStatud.AddItem "Procesando actualizacion en: " & strNombreBaseDatos
    LstStatud.Text = "Procesando actualizacion en: " & strNombreBaseDatos
    LstStatud.Refresh
    RatonNormal
Exit Sub
error_Handler:
    MsgBox Err.Description, vbCritical
    RatonNormal
End Sub

'''Public Sub UPD_Actualizar_SP_FN()
'''Dim AdoDBTXT As ADODB.Recordset
'''
'''Dim IdFile As Long
'''Dim CadenaTime As String
'''Dim Extension As String
'''Dim TextoFile As String
'''Dim FileOrigen As String
'''Dim FileDestino As String
'''Dim Directorio As String
'''Dim Files() As String
'''
'''On Error GoTo error_Handler
'''
'''   'Actualizando archivo de la nueva version del sistemas en las Bases de Datos y los SP, FN
'''    MiTiempo = Time
'''    ProgressBarEstado.Max = 100
'''    ProgressBarEstado.value = 0
'''    CadenaTime = ""
'''
'''    LstStatud.AddItem "Estableciendo conexion al servidor de datos"
'''    LstStatud.Refresh
'''
'''   'Determinar cuales son las tablas fijas que se van a actualizar en un Array
'''    Contador = 0
'''    ProgressBarEstado.value = ProgressBarEstado.value + 1
'''    LstStatud.AddItem "Determinando la informacion para la actualuzacion"
'''    LstStatud.Refresh
'''   'Archivos de Funciones y Procedimientos de la nueva actualizacion
'''    Directorio = Dir(RutaSistema & "\BASES\UPDATE_DB\*.sql", vbNormal) 'Recupera la primera entrada.
'''    Do While Directorio <> ""
'''       If Directorio <> "." And Directorio <> ".." Then
'''          LstStatud.AddItem strNombreBaseDatos & ": Archivo '" & Directorio & "' a subir"
'''          LstStatud.Text = strNombreBaseDatos & ": Archivo '" & Directorio & "' a subir"
'''          LstStatud.Refresh
'''
'''          ReDim Preserve Files(Contador) As String
'''          Files(Contador) = Directorio
'''          Contador = Contador + 1
'''       End If
'''       Directorio = Dir
'''    Loop
'''    ProgressBarEstado.Max = ProgressBarEstado.Max + (Contador * 2)
'''
'''   'Eliminamos y luego creamos la tabla y el SP de Actualizaciones
'''    ProgressBarEstado.value = ProgressBarEstado.value + 1
'''    Ejecutar_SQL_AdoDB "DROP PROCEDURE IF EXISTS sp_Actualizar_SP_FN;"
'''    ProgressBarEstado.value = ProgressBarEstado.value + 1
'''    Ejecutar_SQL_AdoDB "DROP PROCEDURE IF EXISTS sp_Ejecutar_SQL;"
'''    ProgressBarEstado.value = ProgressBarEstado.value + 1
'''    Ejecutar_SQL_AdoDB "DROP TABLE IF EXISTS Actualizacion;"
'''
'''    sSQL = "CREATE TABLE Actualizacion(" _
'''         & "Archivo VARCHAR(100) NULL, " _
'''         & "Extension VARCHAR(3) NULL, " _
'''         & "Documento NVARCHAR(MAX) NULL, " _
'''         & "Con_Item BIT NULL, " _
'''         & "Con_Periodo BIT NULL, " _
'''         & "ID INT IDENTITY NOT NULL PRIMARY KEY);"
'''    Ejecutar_SQL_SP sSQL
'''
'''    ProgressBarEstado.value = ProgressBarEstado.value + 1
'''    Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Ejecutar_SQL.StoredProcedure.sql"), True
'''    ProgressBarEstado.value = ProgressBarEstado.value + 1
'''    Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Actualizar_SP_FN.StoredProcedure.sql"), True
'''
'''   'Empezamos a subir la informacion de los archivos a la TABLA de Actualizacion, para luego ejecutar el SP que empieza la actualizacion
'''    LstStatud.AddItem "Subiendo informacion al servidor"
'''    LstStatud.Refresh
'''
'''    For I = 0 To UBound(Files)
'''        FileOrigen = Files(I)
'''        ProgressBarEstado.value = ProgressBarEstado.value + 1
'''        LstStatud.AddItem strNombreBaseDatos & ": Subiendo datos de " & FileOrigen
'''        LstStatud.Text = strNombreBaseDatos & ": Subiendo datos de " & FileOrigen
'''        LstStatud.Refresh
'''
'''        Extension = RightStrg(FileOrigen, 3)
'''        FileDestino = RutaSistema & "\BASES\UPDATE_DB\" & FileOrigen
'''        TextoFile = Leer_Archivo_Plano(FileDestino)
'''
'''        TextoFile = Replace(TextoFile, vbCr, "[CR]")
'''        TextoFile = Replace(TextoFile, vbLf, "[LF]")
'''        TextoFile = Replace(TextoFile, "'", "[`]")
'''        TextoFile = Replace(TextoFile, "#", "[N]")
'''        TextoFile = Replace(TextoFile, """", "[DC]")
'''
'''        Select Case Extension
'''          Case "sql"
'''               IdFile = InStr(TextoFile, "CREATE")
'''               If IdFile > 0 Then
'''                  TextoFile = MidStrg(TextoFile, IdFile, Len(TextoFile))
'''                  For K = Len(TextoFile) To Len(TextoFile) - 10 Step -1
'''                      If MidStrg(TextoFile, K, 2) = "GO" Then J = K
'''                  Next K
'''                  TextoFile = MidStrg(TextoFile, 1, J - 1)
'''               End If
'''        End Select
'''
'''       'MsgBox FileOrigen & vbCrLf & Len(TextoFile)
'''        FileOrigen = MidStrg(FileOrigen, 1, Len(FileOrigen) - 4)
'''        Ejecutar_SQL_SP "INSERT INTO Actualizacion (Archivo, Extension, Documento) VALUES ('" & FileOrigen & "', '" & Extension & "', '" & TextoFile & "')"
'''    Next I
'''
'''    LstStatud.AddItem "Actualizando caracteres especiales"
'''    LstStatud.Refresh
'''
'''    sSQL = "SELECT Archivo, Extension, Documento " _
'''         & "FROM Actualizacion " _
'''         & "WHERE Extension <> '.' " _
'''         & "ORDER BY Archivo "
'''    Select_AdoDB AdoDBTXT, sSQL
'''    With AdoDBTXT
'''     If .RecordCount > 0 Then
'''         Do While Not .EOF
'''            ProgressBarEstado.value = ProgressBarEstado.value + 1
'''            LstStatud.Text = "Actualizando signos especiales: " & .fields("Archivo")
'''            LstStatud.Refresh
'''            TextoFile = .fields("Documento")
'''            TextoFile = Replace(TextoFile, "[CR]", vbCr)
'''            TextoFile = Replace(TextoFile, "[LF]", vbLf)
'''            TextoFile = Replace(TextoFile, "[`]", "'")
'''            TextoFile = Replace(TextoFile, "[N]", "#")
'''            TextoFile = Replace(TextoFile, "[DC]", """")
'''           .fields("Documento") = TextoFile
'''           .MoveNext
'''         Loop
'''        .UpdateBatch
'''     End If
'''    End With
'''    AdoDBTXT.Close
'''
'''   'Procedemos a ejecutar el SP que actualizara SP, FN y las tablas
'''    LstStatud.AddItem "Procesando actualizacion en: " & strNombreBaseDatos
'''    LstStatud.Text = "Procesando actualizacion en: " & strNombreBaseDatos
'''    LstStatud.Refresh
'''
'''    Actualizar_SP_FN_SP
'''
'''    ProgressBarEstado.value = ProgressBarEstado.Max
'''    CadenaTime = CadenaTime & Format(Time - MiTiempo, FormatoTimes) & vbCrLf
'''    LstStatud.AddItem "FIN DEL PROCESO DE ACTUALIZACION [" & CadenaTime & "]"
'''    LstStatud.Refresh
'''    RatonNormal
'''Exit Sub
'''
'''error_Handler:
'''    MsgBox Err.Description, vbCritical
'''    RatonNormal
'''End Sub
'''---------------------------
'''Public Sub UPD_Actualizar(LstStatud As ListBox, _
'''                          URLinet As Inet, _
'''                          Update_Dir As DirListBox, _
'''                          Update_File As FileListBox, _
'''                          Update_LstTablas As ListBox, _
'''                          Optional Update_Limpiar_Bases As Boolean)
'''Dim AdoAuxDB As ADODB.Recordset
'''Dim AdoCompDB As ADODB.Recordset
'''Dim AdoCompDB1 As ADODB.Recordset
'''Dim AdoDBKeyID As ADODB.Recordset
'''Dim CantCamposUpd As Integer
'''Dim Update_Tipo As Integer
'''Dim Idx As Integer
'''Dim Idy As Integer
'''Dim PIni As Long
'''Dim PFin As Long
'''Dim IJ As Long
'''Dim ContEsp As Long
'''Dim IdTime As Long
'''Dim NumTrans As Long
'''Dim Update_Largo As Long
'''Dim ContProc As Long
'''Dim MaxProc As Long
'''Dim strCnn As String
'''Dim Update_Campo As String
'''Dim Update_Mails As String
'''Dim Update_Excel As String
'''Dim Update_Emp As String
'''Dim TextoBusqueda1 As String
'''Dim ZTablas As String
'''Dim Borrar_Campo As Boolean
'''Dim CambiarTipo As Boolean
'''Dim Crear_Clave_Primaria As Boolean
'''Dim Existe_ID As Boolean
'''
''' 'MsgBox "Presione Aceptar para empezar la actualizacion: " & Periodo_Contable
'''  RatonReloj
'''  ContProc = 0
''' 'Llenamos lista de tablas actuales
'''  UPD_Listar_Tablas Update_LstTablas
'''
'''  MaxProc = (Update_LstTablas.ListCount * 8)
'''
'''  Progreso_Barra.Mensaje_Box = "PROGRESO DEL ACTUALIZACION"
'''  ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''  ConSubDir = False
'''  Contador = 0: FileResp = 0
'''  FechaInicial = "31/12/" & Year(FechaSistema)
'''  sSQL = "DELETE * " _
'''       & "FROM Modulos " _
'''       & "WHERE Modulo = 'VS' "
'''  Ejecutar_SQL_AdoDB sSQL
'''
'''  sSQL = "SELECT MIN(Fecha) As Fecha_MIN " _
'''       & "FROM Facturas " _
'''       & "WHERE Fecha <= #" & BuscarFecha(FechaSistema) & "# "
'''  Select_AdoDB AdoCompDB, sSQL
'''  If AdoCompDB.RecordCount > 0 Then
'''     If Not IsNull(AdoCompDB.fields("Fecha_MIN")) Then FechaInicial = UltimoDiaMes(AdoCompDB.fields("Fecha_MIN"))
'''  End If
'''  AdoCompDB.Close
'''
'''  'Determinar cuales son las tablas fijas que se van a actualizar
'''   Cadena = Dir(RutaSistema & "\BASES\UPDATE_DB\*.DBS", vbNormal) 'Recupera la primera entrada.
'''   ZTablas = ""
'''   Contador = 0
'''   Do While Cadena <> ""
'''      If Cadena <> "." And Cadena <> ".." Then
'''         ZTablas = ZTablas & MidStrg(Cadena, 2, Len(Cadena) - 5) & " | "
'''         Contador = Contador + 1
'''      End If
'''      Cadena = Dir
'''   Loop
'''
'''  'Eliminamos las Funciones y Procedmientos necesarios
'''   Progreso_Barra.Mensaje_Box = "Eliminando FN y SP Principales"
'''   ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''   Eliminar_FN_SP_SQL
'''
''' 'Empezamos a actualizar el programa, primero creamos los SP necesarios para empezar a migrar
'''  Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Ejecutar_SQL.StoredProcedure.sql"), True
'''  Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Eliminar_Indices.StoredProcedure.sql"), True
'''  Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Eliminar_Indices_Temporales.StoredProcedure.sql"), True
'''  Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Eliminar_Tablas_Temporales.StoredProcedure.sql"), True
'''  Ejecutar_SQL_AdoDB Crear_FN_SP(RutaSistema & "\BASES\UPDATE_DB\dbo.sp_Update_Default.StoredProcedure.sql"), True
'''
''' 'Primero eliminamos los indices
'''  Progreso_Barra.Mensaje_Box = "Eliminacion de indices"
'''  ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''  Ejecutar_SP "sp_Eliminar_Indices_Temporales", ""
'''
'''  Progreso_Barra.Mensaje_Box = "Eliminamos actualizacion anterior"
'''  ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
''' 'Eliminar Tablas Temporales que sean necesarias
'''  Ejecutar_SP "sp_Eliminar_Tablas_Temporales", ""
'''
''' Contador = 0
''''TxtResult = ""
'''SubCtaGen = ""
'''TextoBusqueda = ""
'''Cadena = ""
'''Cadena1 = ""
'''
''''Volvemos a actualizar las tablas actuales despues de haber borrado las temporales o vacias
'''Update_Dir.Path = RutaSistema & "\BASES\UPDATE_DB"
'''Update_File.Filename = Update_Dir.Path & "\*.UPD"
'''UPD_Listar_Tablas Update_LstTablas
'''For IJ = 0 To Update_File.ListCount - 1
'''    RutaGeneraFile = RutaSistema & "\BASES\UPDATE_DB\" & Update_File.List(IJ)
'''   'Nombre de la tabla que se va actuaizar
'''    RutaOrigen = TrimStrg(MidStrg(Update_File.List(IJ), 1, Len(Update_File.List(IJ)) - 4))
'''   'Leemos los campos de la tabla
'''    UPD_Leer_Campos_Tabla RutaGeneraFile
'''   'Este procedimiento retorna en la variable "TablaNew" de tipo vector los campos a actualizar
'''    Si_No = False
'''    For I = 0 To Update_LstTablas.ListCount - 1
'''      If RutaOrigen = Update_LstTablas.List(I) Then
'''         Si_No = True
'''         I = Update_LstTablas.ListCount
'''      End If
'''    Next I
'''
'''   'Si la Tabla Existe pasamos a actualizar
'''    If Si_No Then
'''       Progreso_Barra.Mensaje_Box = "Actualizando campos de: " & RutaOrigen
'''       ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''       NombreCampo = ""
'''       Existe_ID = False
'''       Crear_Clave_Primaria = False
'''       CantCamposUpd = CantCampos
'''       sSQL = "SELECT * " _
'''            & "FROM " & RutaOrigen & " " _
'''            & "WHERE 1 = 0 "
'''       'MsgBox RutaOrigen & vbCrLf & vbCrLf & AdoStrCnn
'''       Select_AdoDB AdoCompDB, sSQL
'''       CantCampos = CantCamposUpd
'''       With AdoCompDB
'''            ReDim TablaOld(.fields.Count) As Crear_Tablas
'''            Cadena = ""
'''            Contador = 0
'''           'MsgBox CantCampos
'''            For K = 0 To CantCampos - 1
'''                Progreso_Barra.Mensaje_Box = "Actualizando campos de: " & RutaOrigen & " -> " & TablaNew(K).Campo
'''                Evaluar = False
'''                CambiarTipo = False
'''                For J = 0 To .fields.Count - 1
'''                    If TablaNew(K).Campo = .fields(J).Name Then
'''                       Evaluar = True
'''                       If SQL_Server Then
'''                          If TypeField(TablaNew(K).TipoSQL) <> .fields(J).Type Then CambiarTipo = True
'''                       Else
'''                          If TypeField(TablaNew(K).TipoAccess) <> .fields(J).Type Then CambiarTipo = True
'''                       End If
'''                       If (TablaNew(K).LargoCampo <> 0) And (TablaNew(K).LargoCampo <> .fields(J).DefinedSize) Then CambiarTipo = True
'''                    End If
'''                    If TablaNew(K).Campo = "ID" Then Crear_Clave_Primaria = True
'''                    If .fields(J).Name = "ID" Then Existe_ID = True
'''                Next J
'''               'If RutaOrigen = "Facturas" And TablaNew(K).Campo = "Direccion" Then MsgBox TablaNew(K).Campo
'''                If Evaluar Then
'''                  'Actualizo el campo antiguo si hay cambios
'''                   If CambiarTipo Then
'''                      SQL1 = "ALTER TABLE [" & RutaOrigen & "] "
'''                      If SQL_Server Then
'''                         SQL1 = SQL1 & "ALTER COLUMN [" & TablaNew(K).Campo & "] " & TablaNew(K).TipoSQL
'''                      Else
'''                         SQL1 = SQL1 & "ALTER COLUMN [" & TablaNew(K).Campo & "] " & TablaNew(K).TipoAccess
'''                      End If
'''                      SQL1 = SQL1 & "; "
'''                     'If RutaOrigen = "Trans_Documentos" Then MsgBox "CAMBIAR CAMPO:" & vbCrLf & SQL1
'''                     'Ejecutar_SP "sp_Ejecutar_SQL", SQL1
'''                     Ejecutar_SQL_SP SQL1, True
'''                   End If
'''                Else
'''                  'Si es campo nuevo le actualizo
'''                   SQL1 = ""
'''                   Contador = Contador + 1
'''                   If Contador <= 1 Then Cadena = Cadena & Space(3) & " => "
'''                   Cadena = Cadena & TablaNew(K).Campo & Space(30 - Len(TablaNew(K).Campo))
'''                   If Contador > 3 Then
'''                      Cadena = Cadena & vbCrLf
'''                      Contador = 0
'''                   End If
'''                   If TablaNew(K).Campo <> "ID" Then
'''                      SQL1 = "ALTER TABLE [" & RutaOrigen & "] "
'''                      If SQL_Server Then
'''                         SQL1 = SQL1 & "ADD [" & TablaNew(K).Campo & "] " & TablaNew(K).TipoSQL
'''                      Else
'''                         SQL1 = SQL1 & "ADD [" & TablaNew(K).Campo & "] " & TablaNew(K).TipoAccess
'''                      End If
'''                      SQL1 = SQL1 & "; "
'''                   Else
'''                      If Crear_Clave_Primaria And Not Existe_ID Then
'''                         SQL1 = "ALTER TABLE [" & RutaOrigen & "] "
'''                         If SQL_Server Then
'''                            SQL1 = SQL1 & "ADD [" & TablaNew(K).Campo & "] INT IDENTITY NOT NULL PRIMARY KEY"
'''                         Else
'''                            SQL1 = SQL1 & "ADD [" & TablaNew(K).Campo & "] LONG IDENTITY NOT NULL PRIMARY KEY"
'''                         End If
'''                         SQL1 = SQL1 & "; "
'''                      End If
'''                   End If
'''                   'If Crear_Clave_Primaria And SQL1 <> "" Then MsgBox SQL1
'''                   If TablaNew(K).Campo <> "" And SQL1 <> "" Then Ejecutar_SQL_SP SQL1, True
'''                End If
'''            Next K
'''            If Cadena <> "" Then TextoBusqueda = TextoBusqueda & "Cambio Tabla: " & RutaOrigen & vbCrLf & Cadena & vbCrLf
'''       End With
'''       AdoCompDB.Close
'''    Else
'''      'Si la tabla no existe la creamos
'''       Progreso_Barra.Mensaje_Box = "Creando Tabla nueva: " & RutaOrigen
'''       ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''       TextoBusqueda = TextoBusqueda & "Tabla Nueva: " & RutaOrigen & vbCrLf
'''       Crear_Clave_Primaria = False
'''       For J = 0 To CantCampos - 1
'''           If TablaNew(J).Campo = "ID" Then Crear_Clave_Primaria = True
'''       Next J
'''       SQL1 = "CREATE TABLE [" & RutaOrigen & "] ("
'''       For K = 0 To CantCampos - 1
'''           If TablaNew(K).Campo = "ID" Then
'''              If SQL_Server Then
'''                 SQL1 = SQL1 & "[" & TablaNew(K).Campo & "] INT IDENTITY NOT NULL PRIMARY KEY"
'''              Else
'''                 SQL1 = SQL1 & "[" & TablaNew(K).Campo & "] LONG IDENTITY NOT NULL PRIMARY KEY"
'''              End If
'''           Else
'''              If SQL_Server Then
'''                 SQL1 = SQL1 & "[" & TablaNew(K).Campo & "] " & TablaNew(K).TipoSQL
'''              Else
'''                 SQL1 = SQL1 & "[" & TablaNew(K).Campo & "] " & TablaNew(K).TipoAccess
'''              End If
'''           End If
'''           If K <> (CantCampos - 1) Then SQL1 = SQL1 & ","
'''       Next K
'''       SQL1 = SQL1 & "); "
'''      ' WITH (MEMORY_OPTIMIZED=ON, DURABILITY=SCHEMA_ONLY)
'''      'MsgBox "Creando Tabla Nueva:" & vbCrLf & SQL1
'''       Ejecutar_SQL_SP SQL1, True
'''    End If
'''    RatonNormal
'''Next IJ
'''
''' 'Creamos los SP y FN en sql Server
'''  Progreso_Barra.Mensaje_Box = "Creacion de SP y FN"
'''  ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''  Crear_Script_SQL ProgressBarEstado, LstStatud
'''
''' 'Actualiza el contenido de las tablas con Item 000
'''  ContEsp = Progreso_Barra.Incremento
'''
''' 'Creamos los indices de las tablas
'''  Progreso_Barra.Mensaje_Box = "Creacion de Indices"
'''  ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''  Ejecutar_SP "sp_Eliminar_Indices_Temporales", ""
'''  Ejecutar_SP "sp_Crear_Indices", ""
'''
''' 'Iniciamos datos por defaul
'''  Progreso_Barra.Mensaje_Box = "Iniciar Datos por default"
'''  ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
''''''  Parametros = "'000','.',0"
''''''  'MsgBox "."
''''''  Ejecutar_SP "sp_Iniciar_Datos_Default", Parametros
'''  Iniciar_Datos_Default_SP
'''  'MsgBox "..."
''' '========================================================================
''' 'Actualizamos valores por defecto en los campos con nulos
'''  Cadena = ""
'''  For I = 0 To Update_LstTablas.ListCount - 1
'''     Progreso_Barra.Mensaje_Box = "Actualizando Nulos de " & Update_LstTablas.List(I)
'''     ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''     If MidStrg(Update_LstTablas.List(I), 1, 4) <> "Tipo" Then Eliminar_Nulos_SP Update_LstTablas.List(I)
'''  Next I
'''  If Cadena <> "" Then TextoBusqueda = TextoBusqueda & "Campos con nulos:" & vbCrLf & Cadena
'''  If SubCtaGen <> "" Then TextoBusqueda = TextoBusqueda & SubCtaGen
'''
''' 'Actualizando datos, inserciones y eliminaciones de las tablas de esta actualizacion
'''  UPD_Listar_Tablas Update_LstTablas
'''
''' 'Eliminacion de Tablas que ya no funcionan en la nueva actualizacion
'''  Contador = 0
'''  For I = 0 To Update_LstTablas.ListCount - 1
'''      Si_No = True
'''      For IJ = 0 To Update_File.ListCount - 1
'''          RutaOrigen = TrimStrg(MidStrg(Update_File.List(IJ), 1, Len(Update_File.List(IJ)) - 4))
'''          If Update_LstTablas.List(I) = RutaOrigen Then Si_No = False
'''      Next IJ
'''      If Si_No Then
'''        Progreso_Barra.Mensaje_Box = "Tabla Eliminada: " & Update_LstTablas.List(I)
'''        ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''         If Contador = 0 Then
'''            SubCtaGen = "Tablas Eliminadas: " & vbCrLf & Space(10) & "Tabla: " & Update_LstTablas.List(I) & vbCrLf
'''            Contador = 1
'''         Else
'''            SubCtaGen = SubCtaGen & Space(10) & "Tabla: " & Update_LstTablas.List(I) & vbCrLf
'''         End If
'''        'Eliminar Tabla
'''         SQL1 = "DROP TABLE [" & Update_LstTablas.List(I) & "] "
'''         Ejecutar_SQL_SP SQL1, True
'''         TextoBusqueda = TextoBusqueda & "Tabla: " & Update_LstTablas.List(I) & " fue eliminada " & vbCrLf
'''      End If
'''  Next I
'''
''' '==========================================================================
''' ' Procedemos a crear tablas Temporales de informacion de esta actualizacion
''' '==========================================================================
'''  UPD_Listar_Tablas Update_LstTablas
'''  For IJ = 0 To Update_File.ListCount - 1
'''      RutaGeneraFile = RutaSistema & "\BASES\UPDATE_DB\" & Update_File.List(IJ)
'''     'Nombre de la tabla que se va a crear
'''      RutaOrigen = TrimStrg(MidStrg(Update_File.List(IJ), 1, Len(Update_File.List(IJ)) - 4))
'''     'Leemos los campos de la tabla
'''      UPD_Leer_Campos_Tabla RutaGeneraFile
'''     'Si existe la tala la creamos
'''      If InStr(ZTablas, RutaOrigen) Then
'''        'Creamos la tabla Temporal
'''
'''        Progreso_Barra.Mensaje_Box = "Creando Tabla Temporal: " & RutaOrigen
'''        ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''         Crear_Clave_Primaria = False
'''         For J = 0 To CantCampos - 1
'''             If TablaNew(J).Campo = "ID" Then Crear_Clave_Primaria = True
'''         Next J
'''         SQL1 = "CREATE TABLE [Z" & RutaOrigen & "] ("
'''         For K = 0 To CantCampos - 1
'''             If TablaNew(K).Campo = "ID" Then
'''                If SQL_Server Then
'''                   SQL1 = SQL1 & "[" & TablaNew(K).Campo & "] INT IDENTITY NOT NULL PRIMARY KEY"
'''                Else
'''                   SQL1 = SQL1 & "[" & TablaNew(K).Campo & "] LONG IDENTITY NOT NULL PRIMARY KEY"
'''                End If
'''             Else
'''                If SQL_Server Then
'''                   SQL1 = SQL1 & "[" & TablaNew(K).Campo & "] " & TablaNew(K).TipoSQL
'''                Else
'''                   SQL1 = SQL1 & "[" & TablaNew(K).Campo & "] " & TablaNew(K).TipoAccess
'''                End If
'''             End If
'''             If K <> (CantCampos - 1) Then SQL1 = SQL1 & ","
'''         Next K
'''         SQL1 = SQL1 & "); "
'''         Ejecutar_SQL_SP SQL1, True
'''      End If
'''      RatonNormal
'''  Next IJ
'''
''' 'Subiendo el contenido de la nueva actualizacion
'''
'''    Progreso_Barra.Mensaje_Box = "Subiendo nueva version"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''    UPD_Actualizar_Tablas_Temporales LstStatud
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando nueva version"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''    Ejecutar_SP "sp_UpDate_DB", ""
'''
''' 'Actualizamos los datos de esta version
'''    Progreso_Barra.Mensaje_Box = "Actualizando registro nueva version"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''    Ejecutar_SP "sp_Actualizar_Tablas_Generales", ""
'''
'''    Progreso_Barra.Mensaje_Box = "Actualizando datos por defecto de las nuevas"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''    UPD_Actualizar_Datos_Defecto ProgressBarEstado, LstStatud, URLinet, Update_Dir, Update_File, Update_LstTablas
'''
'''   '========================================================================
'''   'Verificacion de campos que no se pudieron actualizar en la base de datos
'''   '========================================================================
'''    Progreso_Barra.Mensaje_Box = "VERIFICANDO CAMPOS BASES DE DATOS"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''    UPD_Listar_Tablas Update_LstTablas
'''    RatonReloj
'''    Contador = 0
'''    TextoBusqueda1 = ""
'''  For IE = 0 To Update_LstTablas.ListCount - 1
'''      RutaOrigen = Update_LstTablas.List(IE)
'''
'''    Progreso_Barra.Mensaje_Box = "Verificando Tipos Campos: " & RutaOrigen & ", de la tabla"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''      Select Case MidStrg(RutaOrigen, 1, 5)
'''        Case "Tabla", "Tipo_"
'''            Progreso_Barra.Mensaje_Box = "Verificando Tipos Campos: " & RutaOrigen & ", no actualizables"
'''            ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''        Case Else
'''            'Leemos los campos de las tablas de actualizacion con la que esta actualmente
'''             RutaGeneraFile = RutaSistema & "\BASES\UPDATE_DB\" & RutaOrigen & ".Upd"
'''             UPD_Leer_Campos_Tabla RutaGeneraFile
'''             C = CantCampos    ' Cantidad de Campos en la tabla de consulta
'''             sSQL = "SELECT * " _
'''                  & "FROM " & RutaOrigen & " " _
'''                  & "WHERE 1 = 0 "
'''             Select_AdoDB AdoCompDB, sSQL
'''             With AdoCompDB
'''              For JE = 0 To .fields.Count - 1
'''                  Borrar_Campo = True
'''                  For KE = 0 To C - 1
'''                      If .fields(JE).Name = TablaNew(KE).Campo Then
'''                          Borrar_Campo = False
'''                          Update_Largo = TablaNew(KE).LargoCampo
'''                          If SQL_Server Then Update_Campo = TablaNew(KE).TipoSQL Else Update_Campo = TablaNew(KE).TipoAccess
'''                          Update_Tipo = .fields(JE).Type
'''                          If Update_Tipo = 134 Or Update_Tipo = 135 Then Update_Tipo = 7
'''                          If Update_Tipo <> 7 And Update_Largo <> 0 And .fields(JE).DefinedSize <> Update_Largo Then Si_No = True
'''                      End If
'''                  Next KE
'''                 'Borrar el campo que no debe estar en la tabla
'''                  If Borrar_Campo Then
'''
'''                    Progreso_Barra.Mensaje_Box = "Eliminando de " & RutaOrigen & " el Campo: " & .fields(JE).Name
'''                    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''                      SQL1 = "ALTER TABLE " & RutaOrigen & " " _
'''                           & "DROP COLUMN [" & .fields(JE).Name & "];"
'''                      Ejecutar_SQL_SP SQL1
'''                  End If
'''              Next JE
'''             End With
'''             AdoCompDB.Close
'''      End Select
'''  Next IE
'''
''' '========================================================================
'''    Progreso_Barra.Mensaje_Box = "Generando Archivos del Modulo de Auditoria"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''
'''  Update_Excel = ""
''''  TxtResult = ""
'''  RatonNormal
'''  If Cadena <> "" Then TextoBusqueda = TextoBusqueda & "Tablas Actualizadas sin Datos:" & vbCrLf & Cadena
''' ' MsgBox AdoStrCnn
''''  TxtResult.SelStart = Len(TxtResult)
''''  TxtResult.SelLength = Len(TxtResult)
'''''  NombreUsuario = "DiskCover Sytem"
'''  Update_Campo = ""
'''  RutaDestino = ""
'''''  TMail.Adjunto = ""
'''  TMail.MensajeHTML = ""
'''''  EmailEmpresa = ""
'''''  NombreGerente = ""
'''''  Telefono1 = ""
'''''  RazonSocial = ""
'''  ComunicadoEntidad = ""
'''
''''''  sSQL = "SELECT * " _
''''''       & "FROM Empresas " _
''''''       & "ORDER BY Item "
''''''  Select_AdoDB AdoCompDB, sSQL
''''''  With AdoCompDB
''''''   If .RecordCount > 0 Then
''''''       Do While Not .EOF
''''''          If Len(.fields("Razon_Social")) > 1 Then Cadena = .fields("Razon_Social") Else Cadena = .fields("Empresa")
''''''          If Len(.fields("Email")) > 1 And EmailEmpresa = "" Then EmailEmpresa = .fields("Email")
''''''          If Len(.fields("Gerente")) > 1 And NombreGerente = "" Then NombreGerente = .fields("Gerente")
''''''          If Len(.fields("Telefono1")) > 1 And Telefono1 = "" Then Telefono1 = .fields("Telefono1")
''''''          If Len(Cadena) > 1 And RazonSocial = "" Then RazonSocial = Cadena
''''''          TMail.Mensaje = TMail.Mensaje _
''''''                        & "CI/RUC: " & .fields("RUC") & vbTab _
''''''                        & .fields("Ciudad") & vbTab & vbTab _
''''''                        & .fields("Gerente") & vbTab & vbTab _
''''''                        & Cadena & vbCrLf
''''''         .MoveNext
''''''       Loop
''''''   End If
''''''  End With
''''''  AdoCompDB.Close
'''  '---------------------------------------------------------------------------------
'''   If IP_PC.InterNet Then
'''      RatonReloj
'''      sSQL = "SELECT * " _
'''           & "FROM lista_estados " _
'''           & "WHERE Estado <> '.' " _
'''           & "ORDER BY ID,Estado "
'''      Select_AdoDB_MySQL AdoRegMySQL, sSQL
'''      With AdoRegMySQL
'''       If .RecordCount > 0 Then
'''           Do While Not .EOF
'''              sSQL = "SELECT * " _
'''                   & "FROM Tabla_Referenciales_SRI " _
'''                   & "WHERE Tipo_Referencia = 'ESTADO EMPRESA' " _
'''                   & "AND Codigo = '" & .fields("Estado") & "' "
'''              Select_AdoDB AdoReg, sSQL
'''              If AdoReg.RecordCount <= 0 Then
'''                 SQL1 = "INSERT INTO Tabla_Referenciales_SRI (Tipo_Referencia, Codigo, Descripcion) " _
'''                      & "VALUES ('ESTADO EMPRESA', '" & .fields("Estado") & "', '" & .fields("Descripcion") & "');"
'''                 Ejecutar_SQL_SP SQL1
'''              End If
'''              AdoReg.Close
'''             .MoveNext
'''           Loop
'''       End If
'''      End With
'''      AdoRegMySQL.Close
'''   End If
'''
'''   'Codigos Catalogo Ctas_Proceso
'''    Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Ctas_Proceso"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''    Eliminar_Duplicados_SP "Ctas_Proceso", "Periodo,Item,Detalle", "Detalle", "", True
'''
'''   'Codigos Catalogo Seteos_Documentos
'''    Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Seteos Documentos"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''    Eliminar_Duplicados_SP "Seteos_Documentos", "Item, TP, Campo", "TP", "", True
'''
'''   'Eliminar Duplicados en el Catalogo de Cuentas
'''    Progreso_Barra.Mensaje_Box = "Determinando Duplicados de: Catalogo de Cuentas"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''    Eliminar_Duplicados_SP "Catalogo_Cuentas", "Codigo", "", "", True
'''  '---------------------------------------------------------------------------------
'''    ProgressBarEstado.value = Progreso_Barra.Valor_Maximo
'''
'''    Progreso_Barra.Mensaje_Box = "FIN DEL PROCESO DE ACTUALIZACION"
'''    ftp.Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''End Sub

'''Private Sub Importar_Bases_Antiguas()
'''Dim SiID As Boolean
'''Dim SiItem As Boolean
'''Dim SiCod As Boolean
'''
'''Dim ContTAB As Integer
'''
'''Dim NumReg As Long
'''Dim TotalReg As Long
'''
'''Dim CamposFile() As Campos_Tabla
'''
'''Dim NombreTabla As String
'''
'''    Progreso_Barra.Mensaje_Box = "SUBIENDO ABONOS DEL BANCO " & TextoBanco
'''    Progreso_Iniciar
'''
'''    CDialogDir.Filename = RutaSysBases & "\Datos\Total\*.BDD"
'''    CDialogDir.InitDir = RutaSysBases & "\Datos\Total\"
'''    CDialogDir.Flags = cdlOFNFileMustExist + cdlOFNNoChangeDir + cdlOFNHideReadOnly
'''    CDialogDir.Filter = "Archivos BDD|*.BDD"
'''    CDialogDir.DialogTitle = "Abrir Archivo"
'''    CDialogDir.Action = 1
'''    J = InStrRev(CDialogDir.Filename, "\")
'''    If CDialogDir.Filename <> "" And J > 0 Then
'''       File1.Path = MidStrg(CDialogDir.Filename, 1, J)
'''       File1.Pattern = "*.BDD"
'''       Progreso_Barra.Incremento = 0
'''       For I = 0 To File1.ListCount - 1
'''           Progreso_Barra.Valor_Maximo = File1.ListCount
'''           NumReg = 1
'''           TotalReg = 2
'''           NumFile = FreeFile
'''           NombreArchivo = File1.Path & "\" & File1.List(I)
'''           Open NombreArchivo For Input As #NumFile
'''             Do While Not EOF(NumFile)
'''                Line Input #NumFile, Cod_Field
'''                Cod_Field = Replace(Cod_Field, vbCrLf, "")
'''                Select Case NumReg
'''                  Case 1
'''                       NombreTabla = TrimStrg(MidStrg(Cod_Field, InStrRev(Cod_Field, "-") + 1, Len(Cod_Field)))
'''                       Cod_Field = MidStrg(Cod_Field, 1, Len(Cod_Field) - Len(NombreTabla) - 2)
'''                       TotalReg = Val(TrimStrg(MidStrg(Cod_Field, InStrRev(Cod_Field, "-") + 1, Len(Cod_Field))))
'''                       Progreso_Barra.Mensaje_Box = NombreTabla
'''
'''                       If Not Existe_Tabla(NombreTabla) Then GoTo Fin_Tabla
'''                  Case 2
'''                       SiID = False
'''                       SiItem = False
'''                       SiCod = False
'''                       ContTAB = 0
'''                       K = 1
'''                       For J = 1 To Len(Cod_Field)
'''                        If MidStrg(Cod_Field, J, 1) = vbTab Then
'''                           ReDim Preserve CamposFile(ContTAB) As Campos_Tabla
'''                           CamposFile(ContTAB).Campo = MidStrg(Cod_Field, K, J - K)
'''                           Select Case CamposFile(ContTAB).Campo
'''                             Case "ID": SiID = True
'''                             Case "Item": SiItem = True
'''                             Case "Codigo": SiCod = True
'''                           End Select
'''                           K = J + 1
'''                           ContTAB = ContTAB + 1
'''                        End If
'''                       Next J
'''                       Progreso_Barra.Mensaje_Box = "Encerando: " & NombreTabla
'''
'''                       sSQL = "DELETE * " _
'''                            & "FROM " & NombreTabla & " "
'''                       If SiID Then
'''                          sSQL = sSQL & "WHERE ID > 0 "
'''                       ElseIf SiItem Then
'''                          sSQL = sSQL & "WHERE Item <> '.' "
'''                       ElseIf SiCod Then
'''                          sSQL = sSQL & "WHERE Codigo <> 'D' "
'''                       End If
'''                       Ejecutar_SQL_SP sSQL
''''''                       Cadena = ""
''''''                       For J = 0 To UBound(CamposFile)
''''''                           Cadena = Cadena & CamposFile(J).Campo & " = " & CamposFile(J).Valor & vbCrLf
''''''                       Next J
''''''                       MsgBox Cadena
'''                  Case Else
'''                       ContTAB = 0
'''                       K = 1
'''                       For J = 1 To Len(Cod_Field)
'''                        If MidStrg(Cod_Field, J, 1) = vbTab Then
'''                           'MsgBox UBound(CamposFile) & vbCrLf & MidStrg(Cod_Field, K, J - K)
'''                           If ContTAB <= UBound(CamposFile) Then CamposFile(ContTAB).Valor = MidStrg(Cod_Field, K, J - K)
'''                           K = J + 1
'''                           ContTAB = ContTAB + 1
'''                        End If
'''                       Next J
'''                      'Insertamos el registro actual
''''''                       Cadena = ""
'''                       SetAdoAddNew NombreTabla
'''                       For J = 0 To UBound(CamposFile)
'''                           If CamposFile(J).Campo <> "ID" Then SetAdoFields CamposFile(J).Campo, CamposFile(J).Valor
''''''                           Cadena = Cadena & CamposFile(J).Campo & " = " & CamposFile(J).Valor & vbCrLf
'''                       Next J
'''                       SetAdoUpdate
''''''                       MsgBox Cadena
'''                End Select
'''                Progreso_Barra.Mensaje_Box = NombreTabla & ": " & Format(NumReg, "#,##0") & " -> " & Format(TotalReg, "#,##0")
'''
'''                NumReg = NumReg + 1
'''                Contador = Contador + 1
'''             Loop
'''Fin_Tabla:
'''           Close #NumFile
'''       Next I
'''    End If
'''    Progreso_Final
'''End Sub

Public Sub Enviar_Mail_Actualizacion()
   'Enviamos el mail de confirmacion
    sSQL = "SELECT E.Nombre_Entidad, E.Representante, E.RUC_CI_NIC, E.Email_Entidad, L.ID_Empresa " _
         & "FROM entidad As E, lista_empresas As L " _
         & "WHERE L.Base_Datos = '" & strNombreBaseDatos & "' " _
         & "AND E.ID_Empresa = L.ID_Empresa " _
         & "ORDER BY E.ID_Empresa "
    Select_AdoDB_MySQL AdoDBMySQL, sSQL
    If AdoDBMySQL.RecordCount > 0 Then
       RepresentanteEntidad = AdoDBMySQL.Fields("Representante")
       NombreEntidad = AdoDBMySQL.Fields("Nombre_Entidad")
       EmailEmpresa = AdoDBMySQL.Fields("Email_Entidad")
    End If
    AdoDBMySQL.Close
    TMail.Remitente = RepresentanteEntidad
    TMail.para = ""
    Insertar_Mail TMail.para, EmailEmpresa
    Insertar_Mail TMail.para, CorreoUpdate
    TMail.Asunto = "Proceso de actualizacion, exitoso"
    TMail.Mensaje = NombreEntidad & vbCrLf _
                      & "Representante: " & RepresentanteEntidad & vbCrLf _
                      & "Servidor: " & strIPServidor & vbCrLf _
                      & TMail.Mensaje & vbCrLf _
                      & "SERVIRLES ES NUESTRO COMPROMISO, DISFRUTARLO ES EL SUYO."
    FEnviarCorreos.Show 1
End Sub

Public Sub Actualizar_Servidor(Optional Solo_FN_SP As Boolean)
Dim Conn As New ADODB.Connection
Dim MensajeEmail As String
Dim NombreBase() As String
Dim AdoStrCnnTemp As String
Dim ListaBDActualizada As String
Dim PingServer As String
Dim IdBase As Integer
    
    MiTiempo = Time
   'Empezamos a bajar la actualizacion del servidor de las nubes
    Datos_Procesados_BD "Transfiriendo Datos del Servidor..."
    
    Bajar_Archivos_FTP "[1]"

    ListaBDActualizada = ""
    
   'Empezando a realizar los procesos de actualizacion
    FActualizar.Top = 400
    FActualizar.Height = MDI_Y_Max
    FActualizar.Left = MDI_X_Max - FActualizar.width - 200
    If Solo_FN_SP Then
       FrmBaseDatos.BackColor = &H11C0C0
       FrmBaseDatos.ForeColor = &HCC8000
       TxtBaseDatos.BackColor = &HC00000
    Else
       FrmBaseDatos.BackColor = &HC0C000
       FrmBaseDatos.ForeColor = &H800000
       TxtBaseDatos.BackColor = &H404040
    End If
    FrmBaseDatos.Refresh
    TxtBaseDatos.Refresh
    FrmBaseDatos.Top = 2740
    FrmBaseDatos.Height = MDI_Y_Max - 3650
    TxtBaseDatos.Height = FrmBaseDatos.Height - TxtBaseDatos.Top - 100
    FActualizar.Caption = Modulo & ": Actualizando las Bases de [" & strIPServidor & "] "

    RatonReloj
    IdBase = 0
    TxtBaseDatos.Text = ""
    
    Datos_Procesados_BD "Transmitiendo Datos del Servidor: " & vbCrLf & strIPServidor
    ConectarAdodc AdoAux
    Datos_Procesados_BD "Verificando Datos a Transmitir..."
    
    sSQL = "SELECT sys.databases.name,(SUM(sys.master_files.size) * 8/1024) AS size_MB " _
         & "FROM sys.databases JOIN sys.master_files " _
         & "ON sys.databases.database_id = sys.master_files.database_id " _
         & "WHERE sys.databases.name LIKE 'DiskCover%' " _
         & "GROUP BY sys.databases.name " _
         & "ORDER BY sys.databases.name, size_MB; "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         FrmBaseDatos.Caption = "BASE DE DATOS, Tiempo transcurrido: " & Format(Time - MiTiempo, FormatoTimes)
         FrmBaseDatos.Refresh
         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo * .RecordCount
         Progreso_Barra.Mensaje_Box = "PROGRESO DE ACTUALIZACION"
         Datos_Procesados_BD "BASES QUE SE VAN HA ACTUALIZAR:"
         ReDim NombreBase(.RecordCount) As String
         Do While Not .EOF
            NombreBase(IdBase) = .Fields("name")
            Datos_Procesados_BD Format(IdBase, "00") & " [" & NombreBase(IdBase) & "]"
            ListaBDActualizada = ListaBDActualizada & NombreBase(IdBase) & vbCrLf
            IdBase = IdBase + 1
           .MoveNext
         Loop
     End If
    End With
    Set Conn = AdoAux.Recordset.ActiveConnection
    AdoAux.Recordset.Close
    Conn.Close
    
    Datos_Procesados_BD "Empezando la actualizacion..."
    IniIDBase = Val(TxtID.Text)
    FinIDBase = UBound(NombreBase) - 1
    If IniIDBase > FinIDBase Then IniIDBase = FinIDBase
    
   'Empezamos la actualizacion de todo el servidor
    For IdBase = IniIDBase To FinIDBase
        FrmBaseDatos.Caption = "BASE DE DATOS, Tiempo transcurrido: " & Format(Time - MiTiempo, FormatoTimes)
        FrmBaseDatos.Refresh
       'MsgBox "-> " & NombreBase(IdBase) & vbCrLf & "   En proceso..."
        If Ping_IP(strIPServidor) Then
           MiTiempo = Time
           Datos_Procesados_BD Format(IdBase, "00") & " -> " & NombreBase(IdBase) & vbCrLf & "      En proceso..."
           If strNombreBaseDatos <> NombreBase(IdBase) Then
              AdoStrCnn = Replace(AdoStrCnn, strNombreBaseDatos, NombreBase(IdBase))
              strNombreBaseDatos = NombreBase(IdBase)
           End If
           Dolar = 0
           ConSucursal = False
           FActualizar.Caption = Modulo & ": " & strIPServidor
           
           strNombreBaseDatos = NombreBase(IdBase)
           
          'MsgBox AdoStrCnn
             
          'Conectamos la nueva Base de Datos
           ConectarAdodc AdoAux
           ConectarAdodc AdoBusqEmp
           ConectarAdodc AdoEmpresa
            
          'Procedemos a Actualizar la base actual
           UPD_Actualizar_SP
           Datos_Procesados_BD "      [" & Format(Time - MiTiempo, FormatoTimes) & "] Base actualizada con exito."
          
          'Enviamos el mail de confirmacion
           TMail.Mensaje = "Informacion de actualizacion del sistema " & vbCrLf _
                             & "Fecha y Hora [" & FechaSistema & " - " & Format(Time, FormatoTimes) & "]: " & vbCrLf _
                             & "Del Servidor " & strIPServidor & " se ha actualizado la base de Datos: " & NombreBase(IdBase)
           TMail.Remitente = strIPServidor
           TMail.para = CorreoUpdate
           TMail.Asunto = "Actualizacion exitosa de la Base de Datos: " & NombreBase(IdBase) & " (" & Format(IdBase, "00") & ") Duracion: " _
                        & Format(Time - MiTiempo, FormatoTimes)
           FEnviarCorreos.Show 1
        Else
           MsgBox "LA CONEXION NO ESTA ESTABLECIDA" & vbCrLf _
                & "POR FAVOR LLAME AL BENEFICIARIO" & vbCrLf _
                & "PARA QUE CONECTE LA VPN"
           AdoStrCnn = AdoStrCnnTemp
        End If
    Next IdBase
    Datos_Procesados_BD vbCrLf & "ACTUALIZACION DEL SERVIDOR EXITOSA"
    RatonNormal
    FrmBaseDatos.Caption = "BASE DE DATOS, Tiempo transcurrido: " & Format(Time - MiTiempo, FormatoTimes)
    FrmBaseDatos.Refresh
End Sub

'''Public Sub Actualizar_Servidor_SP_FN()
'''Dim Conn As New ADODB.Connection
'''Dim MensajeEmail As String
'''Dim NombreBase() As String
'''Dim AdoStrCnnTemp As String
'''Dim ListaBDActualizada As String
'''Dim PingServer As String
'''Dim IdBase As Integer
'''
'''    MiTiempo = Time
'''   'Empezamos a bajar la actualizacion del servidor de las nubes
'''    Datos_Procesados_BD "Transfiriendo Datos del Servidor..."
'''
'''    Bajar_Archivos_FTP "[4]"
'''
'''    ListaBDActualizada = ""
'''
'''   'Empezando a realizar los procesos de actualizacion
'''    FActualizar.Top = 400
'''    FActualizar.Height = MDI_Y_Max
'''    FActualizar.Left = MDI_X_Max - FActualizar.width - 200
'''    FrmBaseDatos.BackColor = &HC0C000
'''    FrmBaseDatos.ForeColor = &H800000
'''    TxtBaseDatos.BackColor = &HC00000
'''    FrmBaseDatos.Top = 2740
'''    FrmBaseDatos.Height = MDI_Y_Max - 3650
'''    TxtBaseDatos.Height = FrmBaseDatos.Height - TxtBaseDatos.Top - 100
'''    FActualizar.Caption = Modulo & ": Actualizando las Bases de [" & strIPServidor & "] "
'''
'''    RatonReloj
'''    IdBase = 0
'''    TxtBaseDatos.Text = ""
'''
'''    Datos_Procesados_BD "Transmitiendo Datos del Servidor: " & vbCrLf & strIPServidor
'''    ConectarAdodc AdoAux
'''    Datos_Procesados_BD "Verificando Datos a Transmitir..."
'''
'''    sSQL = "SELECT sys.databases.name,(SUM(sys.master_files.size) * 8/1024) AS size_MB " _
'''         & "FROM sys.databases JOIN sys.master_files " _
'''         & "ON sys.databases.database_id = sys.master_files.database_id " _
'''         & "WHERE sys.databases.name LIKE 'DiskCover%' " _
'''         & "GROUP BY sys.databases.name " _
'''         & "ORDER BY sys.databases.name, size_MB; "
'''    Select_Adodc AdoAux, sSQL
'''    With AdoAux.Recordset
'''     If .RecordCount > 0 Then
'''         FrmBaseDatos.Caption = "BASE DE DATOS, Tiempo transcurrido: " & Format(Time - MiTiempo, FormatoTimes)
'''         FrmBaseDatos.Refresh
'''         Progreso_Barra.Valor_Maximo = Progreso_Barra.Valor_Maximo * .RecordCount
'''         Progreso_Barra.Mensaje_Box = "PROGRESO DE ACTUALIZACION"
'''         Datos_Procesados_BD "BASES QUE SE VAN HA ACTUALIZAR:"
'''         ReDim NombreBase(.RecordCount) As String
'''         Do While Not .EOF
'''            NombreBase(IdBase) = .fields("name")
'''            Datos_Procesados_BD Format(IdBase, "00") & " [" & NombreBase(IdBase) & "]"
'''            ListaBDActualizada = ListaBDActualizada & NombreBase(IdBase) & vbCrLf
'''            IdBase = IdBase + 1
'''           .MoveNext
'''         Loop
'''     End If
'''    End With
'''    Set Conn = AdoAux.Recordset.ActiveConnection
'''    AdoAux.Recordset.Close
'''    Conn.Close
'''
'''    Datos_Procesados_BD "Empezando la actualizacion..."
'''    IniIDBase = Val(TxtID.Text)
'''    FinIDBase = UBound(NombreBase) - 1
'''    If IniIDBase > FinIDBase Then IniIDBase = FinIDBase
'''
'''   'Empezamos la actualizacion de todo el servidor
'''    For IdBase = IniIDBase To FinIDBase
'''        FrmBaseDatos.Caption = "BASE DE DATOS, Tiempo transcurrido: " & Format(Time - MiTiempo, FormatoTimes)
'''        FrmBaseDatos.Refresh
'''       'MsgBox "-> " & NombreBase(IdBase) & vbCrLf & "   En proceso..."
'''        If Ping_IP(strIPServidor) Then
'''           MiTiempo = Time
'''           Datos_Procesados_BD Format(IdBase, "00") & " -> " & NombreBase(IdBase) & vbCrLf & "      En proceso..."
'''           If strNombreBaseDatos <> NombreBase(IdBase) Then
'''              AdoStrCnn = Replace(AdoStrCnn, strNombreBaseDatos, NombreBase(IdBase))
'''              strNombreBaseDatos = NombreBase(IdBase)
'''           End If
'''           Dolar = 0
'''           ConSucursal = False
'''           FActualizar.Caption = Modulo & ": " & strIPServidor
'''
'''           strNombreBaseDatos = NombreBase(IdBase)
'''
'''          'MsgBox AdoStrCnn
'''
'''          'Conectamos la nueva Base de Datos
'''           ConectarAdodc AdoAux
'''           ConectarAdodc AdoBusqEmp
'''           ConectarAdodc AdoEmpresa
'''
'''          'Procedemos a Actualizar la base actual
'''           UPD_Actualizar_SP True
'''           Datos_Procesados_BD "      [" & Format(Time - MiTiempo, FormatoTimes) & "] Base actualizada con exito."
'''
'''          'Enviamos el mail de confirmacion
'''           TMail.Mensaje = "Informacion de actualizacion del sistema " & vbCrLf _
'''                         & "Fecha y Hora [" & FechaSistema & " - " & Format(Time, FormatoTimes) & "]: " & vbCrLf _
'''                         & "Del Servidor " & strIPServidor & " se ha actualizado la base de Datos: " & NombreBase(IdBase)
'''           TMail.para = CorreoUpdate
'''           TMail.Asunto = "Actualizacion exitosa de solo los SP y FN en: " & NombreBase(IdBase) & " (" & Format(IdBase, "00") & ") Duracion: " _
'''                        & Format(Time - MiTiempo, FormatoTimes)
'''           FEnviarCorreos.Show 1
'''        Else
'''           MsgBox "LA CONEXION NO ESTA ESTABLECIDA" & vbCrLf _
'''                & "POR FAVOR LLAME AL BENEFICIARIO" & vbCrLf _
'''                & "PARA QUE CONECTE LA VPN"
'''           AdoStrCnn = AdoStrCnnTemp
'''        End If
'''    Next IdBase
'''    Datos_Procesados_BD vbCrLf & "Proceso de Actualizacion del Servidor Exitosa"
'''    RatonNormal
'''    FrmBaseDatos.Caption = "BASE DE DATOS, Tiempo transcurrido: " & Format(Time - MiTiempo, FormatoTimes)
'''    FrmBaseDatos.Refresh
'''End Sub

Public Sub Proceso_Terminado_Exitosamente()
    Progreso_Barra.Incremento = Progreso_Barra.Valor_Maximo
    ProgressBarEstado.Max = Progreso_Barra.Valor_Maximo
    ProgressBarEstado.value = Progreso_Barra.Valor_Maximo
    RatonNormal
   'MsgBox "Proceso realizado con exito"
    LblAdvertencia.ForeColor = &HFFFF80
    LblAdvertencia.FontSize = 18
    LblAdvertencia.FontBold = True
    LblAdvertencia.Caption = "PROCESO DE ACTUALIZACION" & vbCrLf & "FINALIZADO CON EXITO"
End Sub

Public Sub Datos_Procesados_BD(TextoAInsertar As String)
    TxtBaseDatos.Text = TxtBaseDatos.Text & TextoAInsertar & vbCrLf
    TxtBaseDatos.SelStart = Len(TxtBaseDatos.Text)
    TxtBaseDatos.SelLength = Len(TxtBaseDatos.Text)
    TxtBaseDatos.Refresh
End Sub

Private Sub TxtID_GotFocus()
   MarcarTexto TxtID
End Sub

Private Sub TxtID_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtID_LostFocus()
   TxtID.Text = Format(Val(TxtID), "00")
   IniIDBase = Val(TxtID.Text)
End Sub

'''Private Sub Descargar_FTP_Certificados_Logos()
'''Dim AdoDBTemp As ADODB.Recordset
'''Dim ListaDeArchivos As String
'''Dim Certificados() As String
'''Dim LogoTipos() As String
'''Dim Idc As Byte
'''Dim iDL As Byte
'''
'''On Error GoTo error_Handler
'''
'''  Idc = 0
'''  iDL = 0
'''  sSQL = "SELECT Empresa, Ruta_Certificado " _
'''       & "FROM Empresas " _
'''       & "WHERE Ruta_Certificado LIKE '%P12' " _
'''       & "ORDER BY Empresa "
'''  Select_AdoDB AdoDBTemp, sSQL
'''  If AdoDBTemp.RecordCount > 0 Then
'''     Do While Not AdoDBTemp.EOF
'''        RutaDocumentos = RutaSistema & "\CERTIFIC\" & AdoDBTemp.Fields("Ruta_Certificado")
'''        If Len(Dir$(RutaDocumentos)) = 0 Then
'''           ReDim Preserve Certificados(Idc) As String
'''           Certificados(Idc) = AdoDBTemp.Fields("Ruta_Certificado")
'''           Idc = Idc + 1
'''        End If
'''        AdoDBTemp.MoveNext
'''    Loop
'''  End If
'''  AdoDBTemp.Close
'''
'''  sSQL = "SELECT Logo_Tipo " _
'''       & "FROM Empresas " _
'''       & "WHERE Logo_Tipo <> '.' " _
'''       & "GROUP BY Logo_Tipo " _
'''       & "ORDER BY Logo_Tipo; "
'''  Select_AdoDB AdoDBTemp, sSQL
'''  If AdoDBTemp.RecordCount > 0 Then
'''     RatonReloj
'''     Do While Not AdoDBTemp.EOF
'''        RutaDocumentos = RutaSistema & "\LOGOS\" & AdoDBTemp.Fields("Logo_Tipo") & ".jpg"
'''        If Len(Dir$(RutaDocumentos)) = 0 Then
'''           ReDim Preserve LogoTipos(iDL) As String
'''           LogoTipos(iDL) = AdoDBTemp.Fields("Logo_Tipo") & ".jpg"
'''           iDL = iDL + 1
'''        End If
'''        RutaDocumentos = RutaSistema & "\LOGOS\" & AdoDBTemp.Fields("Logo_Tipo") & ".gif"
'''        If Len(Dir$(RutaDocumentos)) = 0 Then
'''           ReDim Preserve LogoTipos(iDL) As String
'''           LogoTipos(iDL) = AdoDBTemp.Fields("Logo_Tipo") & ".gif"
'''           iDL = iDL + 1
'''        End If
'''        RutaDocumentos = RutaSistema & "\LOGOS\" & AdoDBTemp.Fields("Logo_Tipo") & ".png"
'''        If Len(Dir$(RutaDocumentos)) = 0 Then
'''           ReDim Preserve LogoTipos(iDL) As String
'''           LogoTipos(iDL) = AdoDBTemp.Fields("Logo_Tipo") & ".gif"
'''           iDL = iDL + 1
'''        End If
'''        AdoDBTemp.MoveNext
'''    Loop
'''    RatonNormal
'''  End If
'''  AdoDBTemp.Close
'''
'''  With ftp
'''      .Inicializar FActualizar
'''
'''       If EsReadOnly Then
'''          ftpDirUpdate = "/files/" 'ftpUpDir
'''          If InStr(IPDelOrdenador, "192.168.27") Then
'''            .servidor = "192.168.27.3"           'Establecesmo el nombre del Servidor FTP
'''            .Puerto = 21
'''          Else
'''            .servidor = ftpUpSvr                 'Establecesmo el nombre del Servidor FTP
'''            .Puerto = ftpUpPuerto
'''          End If
'''         .Password = ftpUpPwr                    'Le establecemos la contraseña de la cuenta Ftp
'''         .Usuario = ftpUpUse                     'Le establecemos el nombre de usuario de la cuenta
'''       Else
'''         .servidor = ftpSvr                    'Establecesmo el nombre del Servidor FTP
'''         .Password = ftpPwr                    'Le establecemos la contraseña de la cuenta Ftp
'''         .Usuario = ftpUse                     'Le establecemos el nombre de usuario de la cuenta
'''         .Puerto = ftpPuerto
'''          ftpDirUpdate = ""
'''       End If
'''       'MsgBox "..."
''''''      'Le establecemos la contraseña de la cuenta Ftp
''''''      .Password = ftpPwr
''''''      'Le establecemos el nombre de usuario de la cuenta
''''''      .Usuario = ftpUse
''''''      'Establecesmo el nombre del Servidor FTP
''''''       'Or InStr(IP_PC.IP_PC, "192.168.27.") > 0
''''''      'MsgBox IP_PC.IP_PC & vbCrLf &
''''''      'If InStr(IP_PC.IP_PC, "192.168.") > 0 Then .servidor = "192.168.27.4" Else
''''''      .servidor = ftpSvr
''''''      '...conectamos al servidor FTP. EL label es el control donde mostrar los errores y el estado de la conexión
'''       If .ConectarFtp(LstStatud) = False Then
'''           MsgBox "No se pudo conectar"
'''           Exit Sub
'''       End If
'''       LstStatud.Text = LstStatud.Text & .GetDirectorioActual & vbCrLf
'''      .CambiarDirectorio ftpDirUpdate
'''      'Mostramos en el label el path del directorio actual donde estamos ubicados en el servidor
'''      'Le indicamos el ListView donde se listarán los archivos
'''       Set .ListView = LstVwFTP
'''      .ListarArchivos
'''       'MsgBox ftpDirUpdate & vbCrLf & .servidor
'''
'''       If Idc > 0 Then
'''          RatonReloj
'''         'Conectamos la nueva Base de Datos para sacar los Certificados del servidor que no los obtenga el cliente
'''         .CambiarDirectorio ftpDirUpdate & "/SISTEMA/CERTIFIC/"
'''         .ListarArchivos
'''          For I = 1 To LstVwFTP.ListItems.Count
'''              For J = 0 To UBound(Certificados)
'''                  If Certificados(J) = LstVwFTP.ListItems(I) Then
'''                     Progreso_Barra.Mensaje_Box = "Descargando: " & LstVwFTP.ListItems(I)
'''                    .Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''                    .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\CERTIFIC\" & LstVwFTP.ListItems(I), True
'''                     'Exit For
'''                  End If
'''              Next J
'''          Next I
'''          RatonNormal
'''       End If
'''       If iDL > 0 Then
'''          RatonReloj
'''         'Conectamos la nueva Base de Datos para sacar los Certificados del servidor que no los obtenga el cliente
'''         .CambiarDirectorio ftpDirUpdate & "/SISTEMA/LOGOS/"
'''         .ListarArchivos
'''          For I = 1 To LstVwFTP.ListItems.Count
'''              For J = 0 To UBound(LogoTipos)
'''                  If UCaseStrg(LogoTipos(J)) = UCaseStrg(LstVwFTP.ListItems(I)) Then
'''                     Progreso_Barra.Mensaje_Box = "Descargando: " & LstVwFTP.ListItems(I)
'''                    .Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''                    .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\LOGOS\" & LstVwFTP.ListItems(I), True
'''                     'Exit For
'''                  End If
'''              Next J
'''          Next I
'''          RatonNormal
'''       End If
'''
'''       RatonReloj
'''      'Conectamos la nueva Base de Datos para sacar los Certificados del servidor que no los obtenga el cliente
'''      .CambiarDirectorio ftpDirUpdate & "/SISTEMA/FONTSPDF/"
'''      .ListarArchivos
'''       For I = 1 To LstVwFTP.ListItems.Count
'''           Progreso_Barra.Mensaje_Box = "Descargando: " & LstVwFTP.ListItems(I)
'''          .Mostar_Estado_FTP ProgressBarEstado, LstStatud
'''          .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\FONTSPDF\" & LstVwFTP.ListItems(I), True
'''       Next I
'''       RatonNormal
'''
'''      .Desconectar
'''   End With
'''   RatonNormal
'''  'MsgBox "Proceso Terminado con exito"
'''Exit Sub
'''error_Handler:
'''     RatonNormal
'''     MsgBox Err.Description, vbCritical
'''End Sub



