VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form FUpdSystem 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATOS Y PROGRAMAS"
   ClientHeight    =   3015
   ClientLeft      =   105
   ClientTop       =   780
   ClientWidth     =   6540
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   Icon            =   "FUpdSystem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FUpdSystem.frx":1297D
   ScaleHeight     =   3015
   ScaleWidth      =   6540
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImgLstFTP"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir de la actualizacion"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ejecutables"
            Object.ToolTipText     =   "Actualiza solo los ejecutables"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imagenes"
            Object.ToolTipText     =   "Actualizar los Fondos y formatos"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
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
      Top             =   3255
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
      Top             =   3255
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   3885
      TabIndex        =   3
      Top             =   3150
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   3885
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSAdodcLib.Adodc AdoEmpresa 
      Height          =   330
      Left            =   2205
      Top             =   3885
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
      Top             =   4200
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
      Top             =   3255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2205
      Top             =   4200
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
      Top             =   4620
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
      Top             =   3885
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
      Top             =   3255
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
            Picture         =   "FUpdSystem.frx":1C1AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1C4C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1C7E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1CAE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1CE03
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1D11D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1D40F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1DC29
            Key             =   "archivo"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1DF43
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1E25D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1E49B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FUpdSystem.frx":1E7B5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MArchivo 
      Caption         =   "Archivo"
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
Attribute VB_Name = "FUpdSystem"
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

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim hInst As Long
Dim Thread As Long
            
    If Button.key <> "Salir" Then
         RatonReloj
         FUpdSystem.Height = Toolbar1.Top + LstStatud.Top + LstStatud.Height + 850
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
      Case "Ejecutables"                    'Actualiza solo los ejecutables
            Progreso_Barra.Valor_Maximo = 30
            Bajar_Archivos_FTP "[2]"
            TMail.Mensaje = TMail.Mensaje & "Se actualizon Los ejecutables" & vbCrLf
            Enviar_Mail_Actualizacion
            Proceso_Terminado_Exitosamente
      Case "Imagenes"                       'Actualiza solo la imagenes
            Progreso_Barra.Valor_Maximo = 720
            Bajar_Archivos_FTP "[3]"
            TMail.Mensaje = TMail.Mensaje & "Se actualizon los fondos y formatos del sistema" & vbCrLf
            Enviar_Mail_Actualizacion
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
   
   FUpdSystem.Caption = "ESTABLECIENDO CONEXION AL SERVIDOR..."
   FUpdSystem.Refresh
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
         .servidor = ftpSvr                      'Establecesmo el nombre del Servidor FTP
         .Password = ftpPwr                      'Le establecemos la contraseña de la cuenta Ftp
         .Usuario = ftpUse                       'Le establecemos el nombre de usuario de la cuenta
         .Puerto = ftpPuerto
       End If
      'MsgBox .servidor
      'Conectamos al servidor FTP. EL label es el control donde mostrar los errores y el estado de la conexión
       If .ConectarFtp(LstStatud) = False Then
           MsgBox "No se pudo conectar"
           Exit Sub
       End If
       FUpdSystem.Caption = "DATOS Y PROGRAMAS: " & .servidor
       FUpdSystem.Refresh
      'Mostramos en el label el path del directorio actual donde estamos ubicados en el servidor
       Progreso_Barra.Mensaje_Box = .GetDirectorioActual
      'MsgBox Progreso_Barra.Mensaje_Box
      .Mostar_Estado_FTP ProgressBarEstado, LstStatud
      'Le indicamos el ListView donde se listarán los archivos
       Set .ListView = LstVwFTP
       Progreso_Barra.Mensaje_Box = "Buscando directorio en el servidor"
      .Mostar_Estado_FTP ProgressBarEstado, LstStatud
       
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
         
         'Conectamos la nueva Base de Datos para sacar los fondos del Adobe Reader DC
         .CambiarDirectorio "/SISTEMA/FONTSPDF/"
         .ListarArchivos
          For I = 1 To LstVwFTP.ListItems.Count
              If Len(LstVwFTP.ListItems(I)) > 3 Then
                 Progreso_Barra.Mensaje_Box = "Descargando: " & LstVwFTP.ListItems(I)
                .Mostar_Estado_FTP ProgressBarEstado, LstStatud
                .ObtenerArchivo LstVwFTP.ListItems(I), RutaSistema & "\FONTSPDF\" & LstVwFTP.ListItems(I), True
              End If
          Next I
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
      .Desconectar
   End With
   RatonNormal
Exit Sub
error_Handler:
     MsgBox Err.Description, vbCritical
     RatonNormal
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
    CentrarForm FUpdSystem
   '------------------
    Set ping = New cPing
    EsReadOnly = True
    IPDelOrdenador = ping.IP_Del_PC()
   '-----------------------------------
    FUpdSystem.Height = Toolbar1.Top + LstStatud.Top + LstStatud.Height + 820
    
    MDI_X_Max = Screen.width - 150
    MDI_Y_Max = Screen.Height - 1850
    
    
   'Redondear_Cuadro FUpdSystem, 25
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
    FUpdSystem.Caption = Modulo & ": " & strIPServidor & " - " & strNombreBaseDatos
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

'''Private Sub MActOtraBase_Click()
'''  FOtraBase.Top = FUpdSystem.Top + Toolbar1.Height + 150
'''  FOtraBase.Left = FUpdSystem.Left + 150
'''  FOtraBase.Show 1
'''  FUpdSystem.Caption = Modulo & ": " & strIPServidor & " - " & strNombreBaseDatos
'''End Sub

Private Sub MSalir_Click()
   End
End Sub

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

