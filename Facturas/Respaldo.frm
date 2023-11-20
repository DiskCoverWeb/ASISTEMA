VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Respaldos 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Espere un momento....     Estoy procesando las bases"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   Icon            =   "Respaldo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OpcEnviar 
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   945
      TabIndex        =   2
      Top             =   105
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.OptionButton OpcRecibir 
      BackColor       =   &H00FF8080&
      Caption         =   "&Recibir"
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
      Height          =   225
      Left            =   1995
      TabIndex        =   3
      Top             =   105
      Width           =   960
   End
   Begin VB.TextBox TextArchivo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4215
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Text            =   "Respaldo.frx":0442
      Top             =   2310
      Width           =   9045
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFC0C0&
      Height          =   1260
      Left            =   3675
      TabIndex        =   22
      Top             =   945
      Width           =   2535
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
      Height          =   1815
      Left            =   105
      TabIndex        =   1
      Top             =   315
      Width           =   750
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   21
      Top             =   6510
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bajar Enviar &97"
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
      Left            =   9240
      Picture         =   "Respaldo.frx":045B
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "S&ubir Recibir"
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
      Left            =   9240
      Picture         =   "Respaldo.frx":0D25
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3255
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Bajar Enviar"
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
      Left            =   9240
      Picture         =   "Respaldo.frx":15CB
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1155
      Width           =   1065
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
      Height          =   960
      Left            =   9240
      Picture         =   "Respaldo.frx":1DD5
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2205
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   4935
      TabIndex        =   6
      Top             =   315
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
   Begin VB.DriveListBox Drive1 
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
      Height          =   315
      Left            =   945
      TabIndex        =   17
      Top             =   630
      Width           =   2640
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
      Height          =   960
      Left            =   9240
      Picture         =   "Respaldo.frx":20DF
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4305
      Width           =   1065
   End
   Begin VB.DirListBox Dir1 
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
      Height          =   1215
      Left            =   945
      TabIndex        =   18
      Top             =   945
      Width           =   2640
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   420
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
      Left            =   420
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
      Left            =   420
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   3675
      TabIndex        =   5
      Top             =   315
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
   Begin MSForms.CheckBox CheqBNV 
      Height          =   330
      Left            =   6300
      TabIndex        =   24
      Top             =   1680
      Width           =   2115
      VariousPropertyBits=   746588179
      BackColor       =   16744576
      ForeColor       =   65535
      DisplayStyle    =   4
      Size            =   "3731;582"
      Value           =   "0"
      Caption         =   "Nota de Venta"
      PicturePosition =   524294
      Picture         =   "Respaldo.frx":2AD5
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox CheqBFA 
      Height          =   330
      Left            =   6300
      TabIndex        =   23
      Top             =   1365
      Width           =   1590
      VariousPropertyBits=   746588179
      BackColor       =   16744576
      ForeColor       =   65535
      DisplayStyle    =   4
      Size            =   "2805;582"
      Value           =   "0"
      Caption         =   "Facturas"
      PicturePosition =   524294
      Picture         =   "Respaldo.frx":33AF
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox CheqBT 
      Height          =   330
      Left            =   6300
      TabIndex        =   7
      Top             =   105
      Width           =   2850
      VariousPropertyBits=   746588179
      BackColor       =   16744576
      ForeColor       =   65535
      DisplayStyle    =   4
      Size            =   "5027;582"
      Value           =   "0"
      Caption         =   "Todos los Comprobante"
      PicturePosition =   524294
      Picture         =   "Respaldo.frx":3C89
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox CheqBCE 
      Height          =   330
      Left            =   6300
      TabIndex        =   10
      Top             =   1050
      Width           =   2850
      VariousPropertyBits=   746588179
      BackColor       =   16744576
      ForeColor       =   65535
      DisplayStyle    =   4
      Size            =   "5027;582"
      Value           =   "0"
      Caption         =   "Comprobante de Egreso"
      PicturePosition =   524294
      Picture         =   "Respaldo.frx":4233
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox CheqBCI 
      Height          =   330
      Left            =   6300
      TabIndex        =   9
      Top             =   735
      Width           =   2850
      VariousPropertyBits=   746588179
      BackColor       =   16744576
      ForeColor       =   65535
      DisplayStyle    =   4
      Size            =   "5027;582"
      Value           =   "0"
      Caption         =   "Comprobante de Ingreso"
      PicturePosition =   524294
      Picture         =   "Respaldo.frx":48D9
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox CheqBCD 
      Height          =   330
      Left            =   6300
      TabIndex        =   8
      Top             =   420
      Width           =   2745
      VariousPropertyBits=   746588179
      BackColor       =   16744576
      ForeColor       =   65535
      DisplayStyle    =   4
      Size            =   "4842;582"
      Value           =   "0"
      Caption         =   "Comprobante de Diario"
      PicturePosition =   524294
      Picture         =   "Respaldo.frx":4F7F
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &GRUPO"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
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
      Width           =   750
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
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
      Left            =   3675
      TabIndex        =   19
      Top             =   735
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &DESDE:         HASTA:"
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
      Left            =   3675
      TabIndex        =   4
      Top             =   105
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
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
      Left            =   945
      TabIndex        =   16
      Top             =   420
      Width           =   2640
   End
End
Attribute VB_Name = "Respaldos"
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

Public Sub RespaldoFamEnvios(EsAccess97 As Boolean)
  RatonReloj
  If EsAccess97 Then
   ' Procesamos las rutina necesarias antes de respaldar
     PathEmpresa = UCase(RutaEmpresa & "\ENVIOS.MDB")
     AdoStrCnn = XAdoStrCnn & "Data Source=" & PathEmpresa
     ConectarAdodc AdoQuery
  End If
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Beneficiarios..."
  If EsAccess97 Then
     If NumEmpresa = "001" Then
        sSQL = "SELECT DISTINCT 'N' As T,B.*,B.Beneficiario As Cliente,Codigo_B As Codigo,Cod_Ciudad As Grupo,B.CI As FactM " _
             & "FROM Beneficiarios As B,Correos As Co " _
             & "WHERE Co.T = 'P' " _
             & "AND B.Codigo_B = Co.Cod_B " _
             & "ORDER BY Codigo_B "
     Else
        sSQL = "SELECT DISTINCT 'N' As T,B.*,B.Beneficiario As Cliente,Codigo_B As Codigo,Cod_Ciudad As Grupo,B.CI As FactM " _
             & "FROM Beneficiarios As B,Correos As Co " _
             & "WHERE Co.Fecha_P BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND Co.T = 'C' " _
             & "AND B.Codigo_B = Co.Cod_B " _
             & "ORDER BY Codigo_B "
     End If
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Clientes", AdoQuery
  Else
   ' Generamos Tabla:
     Respaldos.Caption = "Tabla: Clientes..."
     If NumEmpresa = "001" Then
        sSQL = "SELECT DISTINCT Cl.*,Cl.Grupo As Cod_Ciudad,FactM As CI,Cl.Codigo As Codigo_B,Cliente As Beneficiario " _
             & "FROM Clientes As Cl," _
             & "Empresas As E,Correos As C " _
             & "WHERE Cl.Codigo = C.Cod_B " _
             & "AND Cl.Grupo = E.Item " _
             & "AND E.Grupo = '" & LGrupo.Text & "' " _
             & "AND C.T = 'P' " _
             & "ORDER BY Codigo "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Clientes", AdoQuery
     Else
        sSQL = "SELECT DISTINCT Cl.*,Cl.Grupo As Cod_Ciudad,FactM As CI,Cl.Codigo As Codigo_B,Cliente As Beneficiario " _
             & "FROM Clientes As Cl," _
             & "Empresas As E,Correos As C " _
             & "WHERE Cl.Grupo = E.Item " _
             & "AND Cl.Codigo = C.Cod_B " _
             & "AND E.Grupo = '" & LGrupo.Text & "' " _
             & "AND C.T = 'C' " _
             & "AND C.Fecha_P BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "ORDER BY Codigo "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Clientes", AdoQuery
     End If
  End If
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Remitentes..."
  If EsAccess97 Then
     If NumEmpresa = "001" Then
        sSQL = "SELECT DISTINCT R.* " _
             & "FROM Remitentes As R,Correos As Co " _
             & "WHERE Co.T = 'P' " _
             & "AND R.Codigo_R = Co.Cod_R " _
             & "ORDER BY Codigo_R "
     Else
        sSQL = "SELECT DISTINCT R.* " _
             & "FROM Remitentes As R,Correos As Co " _
             & "WHERE Co.Fecha_P BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "AND Co.T = 'C' " _
             & "AND R.Codigo_R = Co.Cod_R " _
             & "ORDER BY Codigo_R "
     End If
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Remitentes", AdoQuery
  Else
     If NumEmpresa = "001" Then
        sSQL = "SELECT DISTINCT R.* " _
             & "FROM Remitentes As R," _
             & "Empresas As E," _
             & "Correos AS C " _
             & "WHERE R.Cod_Ciudad = E.Item " _
             & "AND R.Codigo_R = C.Cod_R " _
             & "AND E.Grupo = '" & LGrupo.Text & "' " _
             & "AND C.T = 'P' " _
             & "ORDER BY Codigo_R "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Remitentes", AdoQuery
     Else
        sSQL = "SELECT DISTINCT R.* " _
             & "FROM Remitentes As R," _
             & "Empresas As E," _
             & "Correos As C " _
             & "WHERE R.Cod_Ciudad = E.Item " _
             & "AND R.Codigo_R = C.Cod_R " _
             & "AND E.Grupo = '" & LGrupo.Text & "' " _
             & "AND C.T = 'C' " _
             & "AND C.Fecha_P BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "ORDER BY Codigo_R "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Remitentes", AdoQuery
     End If
  End If
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Corresponsal..."
  sSQL = "SELECT * " _
       & "FROM Corresponsal " _
       & "WHERE Codigo_C <> '.' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Corresponsal", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Corres_Envios..."
  sSQL = "SELECT * " _
       & "FROM Corres_Envios " _
       & "WHERE Codigo_C <> '.' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Corres_Envios", AdoQuery
' Generamos Tabla:
  MiFecha = CLongFecha(CFechaLong(MBFechaI.Text) - 60)
  MiFecha = BuscarFecha(MiFecha)
  If EsAccess97 Then
     If NumEmpresa = "001" Then
        sSQL = "SELECT C.*,Sucursal As SucIng,Sucursal As SucPag " _
             & "FROM Correos As C " _
             & "WHERE Sucursal <> '0' " _
             & "AND T = 'P' " _
             & "ORDER BY Envio_No "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Correos", AdoQuery
        sSQL = "SELECT C.*,Sucursal As SucIng,Sucursal As SucPag " _
             & "FROM Correos As C " _
             & "WHERE Sucursal <> '0' " _
             & "AND C.Fecha BETWEEN #" & MiFecha & "# and #" & FechaFin & "# " _
             & "AND T = 'A' " _
             & "ORDER BY Envio_No "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Correos", AdoQuery
     Else
        sSQL = "SELECT C.*,Sucursal As SucPag,Sucursal As SucIng " _
             & "FROM Correos As C " _
             & "WHERE Sucursal <> '0' " _
             & "AND C.T = 'C' " _
             & "AND C.Fecha_P BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "ORDER BY C.Envio_No "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Correos", AdoQuery
     End If
  Else
     If NumEmpresa = "001" Then
        sSQL = "SELECT C.*,SucIng As Sucursal " _
             & "FROM Correos As C," _
             & "Empresas As E " _
             & "WHERE C.SucIng = E.Item " _
             & "AND E.Grupo = '" & LGrupo.Text & "' " _
             & "AND C.T = 'P' " _
             & "ORDER BY Envio_No "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Correos", AdoQuery
        sSQL = "SELECT C.*,SucIng As Sucursal " _
             & "FROM Correos As C," _
             & "Empresas As E " _
             & "WHERE C.SucIng = E.Item " _
             & "AND C.Fecha BETWEEN #" & MiFecha & "# and #" & FechaFin & "# " _
             & "AND E.Grupo = '" & LGrupo.Text & "' " _
             & "AND C.T = 'A' " _
             & "ORDER BY Envio_No "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Correos", AdoQuery
     Else
        sSQL = "SELECT C.*,SucPag As Sucursal " _
             & "FROM Correos As C," _
             & "Empresas As E " _
             & "WHERE C.SucPag = E.Item " _
             & "AND E.Grupo = '" & LGrupo.Text & "' " _
             & "AND C.T = 'C' " _
             & "AND C.Fecha_P BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
             & "ORDER BY Envio_No "
        SelectData AdoQuery, sSQL
        GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Correos", AdoQuery
     End If
  End If
  If NumEmpresa <> "001" Then
     sSQL = "SELECT '" & GrupoEmpresa & "' As Item,RLL.* " _
          & "FROM Resumen_Llamadas As RLL " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "ORDER BY Fecha,Envio_No "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Resumen_Llamadas", AdoQuery
  End If
  RatonNormal
End Sub

Public Sub RespaldoRetencionesIVA(EsAccess97 As Boolean)
  If EsAccess97 Then
     AdoStrCnn = AdoStrCnnOld
     SQL_Server = Si_No
  End If
  ConectarAdodc AdoQuery
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Clientes", AdoQuery
  If EsAccess97 Then
     sSQL = "SELECT * " _
          & "FROM Trans_Retenciones " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Retenciones", AdoQuery
     sSQL = "SELECT * " _
          & "FROM Trans_Retenciones " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_RolPagos", AdoQuery
     sSQL = "SELECT * " _
          & "FROM Trans_Retenciones " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_RolPagos", AdoQuery
  Else
     sSQL = "SELECT * " _
          & "FROM Trans_Retenciones " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Retenciones", AdoQuery
     sSQL = "SELECT * " _
          & "FROM Trans_RolPagos " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_RolPagos", AdoQuery
     sSQL = "SELECT * " _
          & "FROM Catalogo_RolPagos " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_RolPagos", AdoQuery
  End If
End Sub

Public Sub RespaldoContabilidad(EsAccess97 As Boolean)
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Bancos..."
  If EsAccess97 Then
     sSQL = "SELECT * " _
          & "FROM Bancos " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT * " _
          & "FROM Trans_Bancos " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  If ConSucursal = False Then
     If EsAccess97 Then
        sSQL = sSQL & "AND Item = " & Val(NumEmpresa) & " "
     Else
        sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
     End If
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Bancos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Catalogo_SubCtas..."
  If EsAccess97 Then
     sSQL = "SELECT (TC&TC&Codigo) As Codigo1,'" & NumEmpresa & "' As Item,TC,Beneficiario As Detalle,Presupuesto " _
          & "FROM Beneficiarios " _
          & "WHERE TC = 'I' " _
          & "OR TC = 'G' "
  Else
     sSQL = "SELECT * " _
          & "FROM Catalogo_SubCtas " _
          & "WHERE Item = '" & NumEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_SubCtas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Catalogo_CxCxP..."
  If EsAccess97 Then
     sSQL = "SELECT (TC&TC&Codigo) As Codigo1,'" & GrupoEmpresa & "' As Item,TC,Cta " _
          & "FROM TransaccionesSC " _
          & "WHERE TC = 'C' " _
          & "OR TC = 'P' OR TC = 'R' " _
          & "GROUP BY TC,Codigo,Cta "
  Else
     sSQL = "SELECT * " _
          & "FROM Catalogo_CxCxP " _
          & "WHERE Item = '" & NumEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_CxCxP", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Catalogo_Cuentas..."
  If EsAccess97 Then
     sSQL = "SELECT '" & NumEmpresa & "' As Item,Ct.* " _
          & "FROM Catalogo As Ct " _
          & "WHERE TC <> 'X' " _
          & "ORDER BY Codigo "
  Else
     sSQL = "SELECT * " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "ORDER BY Codigo "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_Cuentas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Comprobantes..."
  If EsAccess97 Then
     sSQL = "SELECT C.*,C.CodigoB As Codigo_B " _
          & "FROM Comprobantes As C " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT * " _
          & "FROM Comprobantes " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  If ConSucursal = False Then
     If EsAccess97 Then
        sSQL = sSQL & "AND Item = " & Val(NumEmpresa) & " "
     Else
        sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
     End If
  End If
  If CheqBT.Value = False Then
     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Comprobantes", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Conciliacion..."
  If EsAccess97 Then
     sSQL = "SELECT '" & NumEmpresa & "' As Item,C.* " _
          & "FROM Conciliacion As C " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT * " _
          & "FROM Trans_Conciliacion " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Conciliacion", AdoQuery
'''' Generamos Tabla:
'''  Respaldos.Caption = "Tabla: Gastos_Caja..."
'''  If EsAccess97 Then
'''     sSQL = "SELECT '" & NumEmpresa & "' As Item,GC.* " _
'''          & "FROM Gastos_Caja As GC " _
'''          & "WHERE TC <> 'X' "
'''  Else
'''     sSQL = "SELECT * " _
'''          & "FROM Catalogo_Gastos_Caja " _
'''          & "WHERE TC <> 'X' "
'''  End If
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_Gastos_Caja", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Presupuestos..."
  If EsAccess97 Then
  sSQL = "SELECT '" & NumEmpresa & "' As Item,P.* " _
       & "FROM Presupuestos As P " _
       & "WHERE Cta <> '.' "
  Else
  sSQL = "SELECT * " _
       & "FROM Trans_Presupuestos " _
       & "WHERE Item = '" & NumEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Presupuestos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Retenciones..."
  If EsAccess97 Then
     sSQL = "SELECT '" & NumEmpresa & "' As Item,R.T,R.TP,R.Numero,R.Cta,R.Fecha," _
          & "(R.Ret_Porc/100) As Porc,R.Valor_Retenido As Valor_Ret,'303' As TD," _
          & "C.CodigoB As Codigo,'RF' As CodigoTR,Retencion As Retencion_No,R.Valor_Factura As Valor_Fact " _
          & "FROM Retenciones As R,Comprobantes As C " _
          & "WHERE R.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND R.TP = C.TP " _
          & "AND R.Numero = C.Numero "
  Else
     sSQL = "SELECT * " _
          & "FROM Trans_Retenciones " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  If ConSucursal = False Then
     If EsAccess97 Then
        sSQL = sSQL & "AND R.Item = " & Val(NumEmpresa) & " " _
             & "ORDER BY R.TP,R.Numero,R.Cta,R.Fecha "
     Else
        sSQL = sSQL & "AND Item = '" & NumEmpresa & "' " _
             & "ORDER BY TP,Numero,Cta,Fecha "
     End If
  End If
  If CheqBT.Value = False Then
     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
  End If

  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Retenciones", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Transacciones..."
  sSQL = "SELECT * " _
       & "FROM Transacciones " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  If ConSucursal = False Then
     If EsAccess97 Then
        sSQL = sSQL & "AND Item = " & Val(NumEmpresa) & " "
     Else
        sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
     End If
  End If
  If CheqBT.Value = False Then
     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
  End If

  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Transacciones", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: TransaccionesGC..."
  If EsAccess97 Then
     sSQL = "SELECT * " _
          & "FROM TransaccionesGC " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT * " _
          & "FROM Trans_Gastos_Caja " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  If ConSucursal = False Then
     If EsAccess97 Then
        sSQL = sSQL & "AND Item = " & Val(NumEmpresa) & " "
     Else
        sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
     End If
  End If

  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Gastos_Caja", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: TransaccionesSC..."
  If EsAccess97 Then
     sSQL = "SELECT (TC&TC&Codigo) As Codigo1,C.* " _
          & "FROM TransaccionesSC As C " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT * " _
          & "FROM Trans_SubCtas " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  End If
  If ConSucursal = False Then
     If EsAccess97 Then
        sSQL = sSQL & "AND Item = " & Val(NumEmpresa) & " "
     Else
        sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
     End If
  End If
  If CheqBT.Value = False Then
     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
  End If
  
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_SubCtas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Kardex..."
  If EsAccess97 Then
     sSQL = "SELECT 'PP' & Codigo_P As Codigo_P1,K.*,TP As TC,Cta As Contra_Cta " _
          & "FROM Kardex As K " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT * " _
          & "FROM Trans_Kardex " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
  End If
  If CheqBT.Value = False Then
     If CheqBCD.Value = False Then sSQL = sSQL & "AND TP <> 'CD' "
     If CheqBCI.Value = False Then sSQL = sSQL & "AND TP <> 'CI' "
     If CheqBCE.Value = False Then sSQL = sSQL & "AND TP <> 'CE' "
     If CheqBFA.Value = False Then sSQL = sSQL & "AND TP <> 'FA' "
     If CheqBNV.Value = False Then sSQL = sSQL & "AND TP <> 'NV' "
  End If
  
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Kardex", AdoQuery
End Sub

Public Sub RespaldoFacturacion(EsAccess97 As Boolean)
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Detalle_Factura..."
  If EsAccess97 Then
     sSQL = "SELECT '" & NumEmpresa & "' As Item,('FA' & Codigo_C) As CodigoC,DF.*,'FA' As TC,Factura_No As Factura " _
          & "FROM Detalle_Factura As DF " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT * " _
          & "FROM Detalle_Factura " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Detalle_Factura", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Facturas..."
  If EsAccess97 Then
     sSQL = "SELECT '" & NumEmpresa & "' As Item,('FA' & Codigo_C) As CodigoC,'FA' As TC,F.* " _
          & "FROM Facturas As F " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT * " _
          & "FROM Facturas " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Facturas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Abonos..."
  If EsAccess97 Then
     sSQL = "SELECT T,'" & NumEmpresa & "' As Item,'FA' & Codigo_C As CodigoC,CtaxCob As Cta," _
          & "CtaxCob As Cta_CxP,TP,Fecha,Diario_No As Recibo_No,Factura,Abonos_MN As Abono " _
          & "FROM Diario_Caja " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND TP = 'CxC' "
  Else
     sSQL = "SELECT * " _
          & "FROM Trans_Abonos " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Abonos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Suscripciones..."
  If EsAccess97 Then
'''     sSQL = "SELECT T,'" & NumEmpresa & "' As Item,'FA' & Codigo_C As Cuenta_No," _
'''          & "Area As Cta,TP,Contrato_No As Credito_No,S.Contador As Tasa," _
'''          & "S.Desde As Fecha,S.Hasta As Fecha_C " _
'''          & "FROM Contratos_Suscrip As S " _
'''          & "WHERE Contrato_No <> '.' "
'''     SelectData AdoQuery, sSQL
'''     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Prestamos", AdoQuery
  Else
     sSQL = "SELECT * " _
          & "FROM Prestamos " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Item = '" & NumEmpresa & "' "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Prestamos", AdoQuery
  End If
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Productos..."
  If EsAccess97 Then
     sSQL = "SELECT '" & NumEmpresa & "' As Item,Ct.*," _
          & "TP As TC,Cta_Inv As Cta_Inventario,Cta As Cta_Proveedor," _
          & "Cta1 As Cta_Costo_Venta,Cta_Ingreso As Cta_Ventas " _
          & "FROM Productos As Ct " _
          & "WHERE TP <> 'X' " _
          & "ORDER BY Codigo_Inv "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_Productos", AdoQuery
  End If
  If EsAccess97 Then
     sSQL = "SELECT '" & NumEmpresa & "' As Item,Codigo As Codigo_Inv,Articulo As Producto," _
          & "'P' As TC,Cta_Inv As Cta_Inventario,Cta_CxP As Cta_Proveedor," _
          & "Cta_Costo As Cta_Costo_Venta,Cta_Ingreso As Cta_Ventas " _
          & "FROM Articulo As Ct " _
          & "WHERE CodigoL <> '.' " _
          & "ORDER BY Codigo "
  Else
     sSQL = "SELECT * " _
          & "FROM Catalogo_Productos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "ORDER BY Codigo_Inv "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Catalogo_Productos", AdoQuery
  If EsAccess97 = False Then
     sSQL = "SELECT TS.*,P.Credito_No " _
          & "FROM Trans_Suscripciones As TS,Prestamos As P " _
          & "WHERE TS.Item = '" & NumEmpresa & "' " _
          & "AND P.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND TS.Contrato_No = P.Credito_No " _
          & "ORDER BY TS.Contrato_No,TS.Fecha "
     SelectData AdoQuery, sSQL
     GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Suscripciones", AdoQuery
  End If
End Sub

Public Sub RespaldoCajaCredito(EsAccess97 As Boolean)
' Procesamos las rutina necesarias antes de respaldar
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Abono_De_Prestamo..."
  sSQL = "SELECT * " _
       & "FROM Abono_De_Prestamo " _
       & "WHERE DC <> 'X' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Abono_De_Prestamo", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Asiento_De_Prestamo..."
  sSQL = "SELECT * " _
       & "FROM Asiento_De_Prestamo " _
       & "WHERE DC <> 'X' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Asiento_De_Prestamo", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Bloqueos..."
  sSQL = "SELECT * " _
       & "FROM Trans_Bloqueos " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Bloqueos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Bancos..."
  sSQL = "SELECT * " _
       & "FROM Trans_Dep_Chq " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Dep_Chq", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Cobranzas..."
  sSQL = "SELECT * " _
       & "FROM Cobranzas " _
       & "WHERE Porc_C <> 0 "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Cobranzas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Conyugue..."
  sSQL = "SELECT * " _
       & "FROM Conyugue " _
       & "WHERE Codigo <> '.' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Conyugue", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Cuentas..."
  sSQL = "SELECT * " _
       & "FROM Cuentas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Cuentas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Garantes..."
  sSQL = "SELECT * " _
       & "FROM Garantes " _
       & "WHERE TP <> 'X' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Garantes", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Socios..."
  sSQL = "SELECT * " _
       & "FROM Socios " _
       & "WHERE CI <> '.' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Socios", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Monto_Apertura..."
  sSQL = "SELECT * " _
       & "FROM Monto_Apertura " _
       & "WHERE Monto_Aper<>0 "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Monto_Apertura", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: MontoA_Tarjeta..."
  sSQL = "SELECT * " _
       & "FROM MontoA_Tarjeta " _
       & "WHERE Monto_Aper<>0 "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "MontoA_Tarjeta", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Prestamos..."
  sSQL = "SELECT P.*,'" & NumEmpresa & "' As Item " _
       & "FROM Prestamos As P " _
       & "WHERE P.Fecha >= #" & FechaIni & "# "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Prestamos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Saldo_Caja_Libreta..."
  sSQL = "SELECT * " _
       & "FROM Saldo_Caja_Libreta " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Saldo_Caja_Libreta", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Saldo_Libretas_Intereses..."
'''  If EsAccess97 Then
'''     sSQL = "SELECT * " _
'''          & "FROM Saldo_Libretas " _
'''          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''  Else
'''     sSQL = "SELECT * " _
'''          & "FROM Saldo_Libretas_Intereses " _
'''          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
'''  End If
'''  SelectData AdoQuery, sSQL
'''  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Saldo_Libretas_Intereses", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Tasa_Interes..."
  sSQL = "SELECT * " _
       & "FROM Tasa_Interes " _
       & "WHERE Desde>=0 "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Tasa_Interes", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Tipo_Prestamo..."
  If EsAccess97 Then
     sSQL = "SELECT '" & NumEmpresa & "' As Item," _
          & "TP.* " _
          & "FROM Tipo_Prestamo As TP " _
          & "WHERE TP <> '.' "
  Else
     sSQL = "SELECT * " _
          & "FROM Tipo_Prestamo " _
          & "WHERE TP <> '.' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Tipo_Prestamo", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Tipo_Proceso..."
  sSQL = "SELECT * " _
       & "FROM Tipo_Proceso " _
       & "WHERE DC <> '.' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Tipo_Proceso", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Cajas..."
  sSQL = "SELECT * " _
       & "FROM Trans_Cajas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  If EsAccess97 Then
     sSQL = sSQL & "AND Item = " & Val(NumEmpresa) & " "
  Else
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Cajas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Libretas..."
  sSQL = "SELECT * " _
       & "FROM Trans_Libretas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  If EsAccess97 Then
     sSQL = sSQL & "AND Item = " & Val(NumEmpresa) & " "
  Else
     sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Libretas", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Prestamos..."
  sSQL = "SELECT TP.*,'" & NumEmpresa & "' As Item " _
       & "FROM Trans_Prestamos As TP " _
       & "WHERE TP.Fecha >= #" & FechaIni & "# "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_Prestamos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Trans_Prestamos..."
  sSQL = "SELECT TP.*,'" & NumEmpresa & "' As Item " _
       & "FROM Trans_Prestamos As TP " _
       & "WHERE TP.Fecha_C " _
       & "BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND TP.T = 'C' "
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Trans_PrestamosC", AdoQuery
  ' Generamos Tabla:
  Respaldos.Caption = "Tabla: Clientes..."
  If EsAccess97 Then
     sSQL = "SELECT '" & NumEmpresa & "' As Grupo,C.*,(Apellidos & ' ' & Nombres) As Cliente " _
          & "FROM Cuentas As C " _
          & "WHERE C.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND C.Item = " & Val(NumEmpresa) & " "
  Else
     sSQL = "SELECT * " _
          & "FROM Clientes " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "AND Grupo = '" & NumEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Clientes", AdoQuery
End Sub

Public Sub TipoRespaldo(EsAccess97 As Boolean)
  RatonReloj
  FAConLineas = True
  GrupoEmpresa = LGrupo.Text
  Contador = 0: FileResp = 0
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
' Eliminamos archivos de otros dias
  'MsgBox Dir1.Path
  File1.FileName = Dir1.Path & "\F*.TXT"
  File1.Refresh
  If File1.ListCount > 0 Then Kill Dir1.Path & "\F*.TXT"
  If EsAccess97 Then
     SQL_ServerOld = SQL_Server
     AdoStrCnnOld = AdoStrCnn
     SQL_Server = False
   ' Buscamos la cadena de conección a la base
     RutaGeneraFile = RutaSistema & "\CONECTAR.TXT"
     NumFile = FreeFile
     AdoStrCnn = ""
     Open RutaGeneraFile For Input As #NumFile
     Do While Not EOF(NumFile)
        AdoStrCnn = AdoStrCnn & Input(1, #NumFile) ' Obtiene un carácter.
     Loop
     Close #NumFile
     XAdoStrCnn = AdoStrCnn
   ' Procesamos las rutina necesarias antes de respaldar
     PathEmpresa = UCase(RutaSistema & "\EMPRESAS.MDB")
     AdoStrCnn = XAdoStrCnn & "Data Source=" & PathEmpresa
     ConectarAdodc AdoQuery
     sSQL = "SELECT * " _
          & "FROM Empresas " _
          & "WHERE Item = " & Val(GrupoEmpresa) & " "
     SelectData AdoQuery, sSQL
     If AdoQuery.Recordset.RecordCount > 0 Then Carpeta = AdoQuery.Recordset.Fields("SubDir")
     PathEmpresa = UCase(RutaSistema & "\BASES97\" & UCase(Carpeta) & ".MDB")
     AdoStrCnn = XAdoStrCnn & "Data Source = " & PathEmpresa
     'MsgBox AdoStrCnn
     ConectarAdodc AdoAux
     ConectarAdodc AdoAct
     ConectarAdodc AdoQuery
  End If
 'Preparamos los codigos de Clientes, Proveedores y suscriptores
  If EsAccess97 Then PrepararClientes97
  
  TextArchivo.Text = "TABLAS PROCESADAS:" & vbCrLf _
                   & "==================" & vbCrLf
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Empresas..."
  If EsAccess97 Then
     sSQL = "SELECT * " _
          & "FROM Empresas " _
          & "WHERE Item = " & Val(GrupoEmpresa) & " " _
          & "ORDER BY Item "
  Else
     sSQL = "SELECT * " _
          & "FROM Empresas " _
          & "WHERE Item = '" & GrupoEmpresa & "' " _
          & "ORDER BY Item "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Empresas", AdoQuery
' Generamos Tabla:
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
       TextoFileEmp = "10 - Fecha_Respaldo" & vbCrLf _
                    & "Datos_Respaldo|" & vbCrLf _
                    & "  1.- Empresa/Sucursal: " & .Fields("Empresa") & vbCrLf _
                    & "  2.- R.U.C.          : " & .Fields("RUC") & vbCrLf _
                    & "  3.- Gerente         : " & .Fields("Gerente") & vbCrLf _
                    & "  4.- Telefono        : " & .Fields("Telefono1") & vbCrLf _
                    & "  5.- FAX             : " & .Fields("FAX") & vbCrLf _
                    & "  6.- Numero Asignado : " & .Fields("Item") & vbCrLf _
                    & "  7.- Fecha Inicial   : " & MBFechaI.Text & " (" & FechaDiaSem(MBFechaI.Text) & ")" & vbCrLf _
                    & "  8.- Fecha Final     : " & MBFechaF.Text & " (" & FechaDiaSem(MBFechaF.Text) & ")" & vbCrLf _
                    & "  9.- Modulo          : " & UCase(Modulo) & vbCrLf _
                    & " 10.- Usuario         : " & NombreUsuario & vbCrLf _
                    & String(55, "_") & vbCrLf _
                    & "ARCHIVOS PROCESADOS:" & vbCrLf _
                    & "====================" & vbCrLf _
                    & "F" & Format(Day(MBFechaI.Text), "00") _
                    & Format(Month(MBFechaI.Text), "00") _
                    & "000.TXT" & " => Fecha_Respaldo"
   End If
  End With
  RutaEmpresa = UCase(RutaSistema & "\EMPRESA\" & Carpeta)
  RutaEmpresaOld = UCase(RutaSistema & "\EMPRESA\" & Carpeta)
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Accesos..."
  If EsAccess97 Then
     sSQL = "SELECT '" & GrupoEmpresa & "' As Item,A.* " _
          & "FROM Accesos As A " _
          & "WHERE Mid(Codigo,1,6) <> 'ACCESO' " _
          & "ORDER BY Codigo "
  Else
     sSQL = "SELECT * " _
          & "FROM Accesos " _
          & "WHERE Item = '" & GrupoEmpresa & "' " _
          & "ORDER BY Codigo "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Accesos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Acceso_Empresa..."
  If EsAccess97 Then
     sSQL = "SELECT '" & GrupoEmpresa & "' As Item,'99' As Modulo,AE.* " _
          & "FROM Acceso_Empresa As AE " _
          & "WHERE Codigo <> '.' " _
          & "ORDER BY Codigo "
  Else
     sSQL = "SELECT * " _
          & "FROM Acceso_Empresa " _
          & "WHERE Codigo <> '.' " _
          & "ORDER BY Codigo "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Acceso_Empresa", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: Codigos..."
  If EsAccess97 Then
     sSQL = "SELECT '" & GrupoEmpresa & "' As Item,C.* " _
          & "FROM Codigos As C " _
          & "WHERE Numero <> 0 " _
          & "ORDER BY Concepto "
  Else
     sSQL = "SELECT * " _
          & "FROM Codigos " _
          & "WHERE Item = '" & GrupoEmpresa & "' " _
          & "ORDER BY Concepto "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Codigos", AdoQuery
' Generamos Tabla:
  Respaldos.Caption = "Tabla: CtasProceso..."
  If EsAccess97 Then
     sSQL = "SELECT '" & GrupoEmpresa & "' As Item,CP.*,Cuenta As Detalle " _
          & "FROM CtasProceso As CP " _
          & "WHERE Codigo <> '.' "
  Else
     sSQL = "SELECT * " _
          & "FROM Ctas_Proceso " _
          & "WHERE Item = '" & GrupoEmpresa & "' "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Ctas_Proceso", AdoQuery
 'Generamos Tabla:
  Respaldos.Caption = "Tabla: Clientes..."
  If EsAccess97 Then
     sSQL = "SELECT * " _
          & "FROM Clientes " _
          & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# "
  Else
     sSQL = "SELECT DISTINCT C.*,Co.Codigo_B As CodigoAux " _
          & "FROM Clientes As C,Comprobantes AS Co " _
          & "WHERE Co.Item = '" & GrupoEmpresa & "' " _
          & "AND C.Codigo = Co.Codigo_B " _
          & "UNION " _
          & "SELECT DISTINCT C.*,F.CodigoC As CodigoAux " _
          & "FROM Clientes As C,Facturas As F " _
          & "WHERE F.Item = '" & GrupoEmpresa & "' " _
          & "AND C.Codigo = F.CodigoC " _
          & "UNION " _
          & "SELECT DISTINCT C.*,P.Cuenta_No As CodigoAux " _
          & "FROM Clientes As C,Prestamos As P " _
          & "WHERE P.Item = '" & GrupoEmpresa & "' " _
          & "AND C.Codigo = P.Cuenta_No " _
          & "ORDER BY C.Cliente "
  End If
  SelectData AdoQuery, sSQL
  GenerarTablaArchivoPlano Respaldos, MBFechaI.Text, "Clientes", AdoQuery
  RespaldoContabilidad EsAccess97
'   RespaldoFacturacion EsAccess97

''''  RespaldoCajaCredito EsAccess97
''''  RespaldoFamEnvios EsAccess97
''''  RespaldoRetencionesIVA EsAccess97
  ' Generamos Tabla:
  Codigo = "F" & Format(Day(MBFechaI.Text), "00") & Format(Month(MBFechaI.Text), "00") & "000.TXT"
  RutaGeneraFile = Left(Drive1.Drive, 2) & "\SYSBASES\DATOS\E" & GrupoEmpresa & "\" & Codigo
  TextArchivo.Text = TextArchivo.Text & Space(7) & Codigo & " => Fecha_Respaldo" & vbCrLf
  TextArchivo.Refresh
  NumFile = FreeFile
  Open RutaGeneraFile For Output As #NumFile ' Abre el archivo.
  FAConLineas = True
  Print #NumFile, TextoFileEmp
  Print #NumFile, String(55, "_")
  Close #NumFile
' Restauramos la cadena de conexion
  If EsAccess97 Then
     AdoStrCnn = AdoStrCnnOld
     SQL_Server = SQL_ServerOld
     ConectarAdodc AdoAux
     ConectarAdodc AdoAct
     ConectarAdodc AdoQuery
  End If
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  Respaldos.Caption = "MODULO DE RESPALDOS"
  FUnidad.Show 1
  If RutaSubDirTemp <> "" Then
' MsgBox "Preceso Terminado"
  Codigo = Mid(RutaDestino, 1, 2)
  If (Codigo = "A:") Or (Codigo = "B:") Then
     RutaOrigen = Codigo & "\Z" _
                & Format(Day(MBFechaI.Text), "00") _
                & Format(Month(MBFechaI.Text), "00") _
                & GrupoEmpresa & ".ZIP"
     RutaDestino = Dir1.Path & "\*.TXT"
  Else
     RutaOrigen = Dir1.Path & "\Z" _
                & Format(Day(MBFechaI.Text), "00") _
                & Format(Month(MBFechaI.Text), "00") _
                & GrupoEmpresa & ".ZIP"

     RutaDestino = Dir1.Path & "\F" _
                 & Format(Day(MBFechaI.Text), "00") _
                 & Format(Month(MBFechaI.Text), "00") _
                 & "*.TXT"
     Cadena = Dir(RutaOrigen, vbArchive)
     If Cadena <> "" Then Kill RutaOrigen
  End If
  MsgBox "Se generara el archivo:" & vbCrLf & vbCrLf & RutaOrigen & vbCrLf & RutaDestino
  Cadena = "SQLRESPA.BAT " _
         & Codigo & " " _
         & RutaOrigen & " " _
         & RutaDestino
  Shell Cadena, vbMaximizedFocus
  End If
End Sub

Public Sub AbrirCamposSQL(NumFile As Integer)
    Cod_Emp = "": Cod_Base = "": Cod_Field = ""
    Line Input #NumFile, Cod_Base
    TotalReg = CLng(SinEspaciosIzq(Cod_Base))
    Cod_Base = SinEspaciosDer(Cod_Base)
    Line Input #NumFile, Cod_Field
    'MsgBox Cod_Base & vbCrLf & Cod_Field
    CantCampos = 0
    For I = 1 To Len(Cod_Field)
        If Mid(Cod_Field, I, 1) = "|" Then CantCampos = CantCampos + 1
    Next I
    ReDim TipoC(CantCampos) As Campos_Tabla
    No_Desde = 1: No_Hasta = 1
    Cadena = Cod_Field
    For I = 1 To CantCampos
        Do
           No_Hasta = No_Hasta + 1
        Loop Until Mid(Cadena, No_Hasta, 1) = "|"
        TipoC(I).Campo = Trim(Mid(Cadena, No_Desde, No_Hasta - 1))
        Cadena = Mid(Cadena, No_Hasta + 1, Len(Cadena))
        No_Desde = 1: No_Hasta = 1
    Next I
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

Public Sub EjecutarRespaldo()
  RatonReloj
  File1.FileName = Dir1.Path & "\*.ZIP"
  File2.FileName = Dir1.Path & "\*.TXT"
  If File2.ListCount > 0 Then Kill RutaBackup & "\*.TXT"
  TextUnidad.Text = UCase(TextUnidad.Text)
  Respaldos.Caption = "Restaurando las bases "
  If (TextUnidad.Text & ":" = "A:") Or (TextUnidad.Text & ":" = "B:") Then
     ChDrive TextUnidad.Text & ":"
     Shell "restaura.bat " & TextUnidad.Text & ": " & Mid(RutaSistema, 1, 2) & " " & File1.FileName, vbMaximizedFocus
  Else
     ChDrive TextUnidad.Text & ":"
     ChDir TextUnidad.Text & ":\SYSBASES"
     Shell "restaura.bat " & TextUnidad.Text & ": Ninguno " & File1.FileName, vbMaximizedFocus
  End If
  ChDrive Mid(RutaSistema, 1, 2)
  Respaldos.Caption = "RESPALDOS DE BASES"
  File1.FileName = Dir1.Path & "\*.ZIP"
  File2.FileName = Dir1.Path & "\*.TXT"
  RatonNormal
End Sub

Public Sub AbrirArchivoSQL(NumFile As Integer)
    RatonReloj
    TextArchivo.Text = ""
    Cod_Emp = ""
    Cod_Base = ""
    Cod_Field = ""
    Cod_NumEmp = ""
    Cod_FechaI = ""
    Cod_FechaF = ""
    Do While Not EOF(NumFile)
       Line Input #NumFile, Cadena
       TextArchivo.Text = TextArchivo.Text & Cadena & vbCrLf
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
  RatonReloj
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  NombreTabla = Trim(NombreTabla)
  Cont1 = 0
  sSQL = "DELETE * " _
       & "FROM " & NombreTabla & " " _
       & "WHERE Fecha " _
       & "BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  ConectarAdoExecute sSQL
  
  'MsgBox NombreTabla & vbCrLf & FechaIni & vbCrLf & FechaFin & vbCrLf & sSQL
  
  sSQL = "SELECT * FROM " & NombreTabla & " " _
       & "WHERE Item = '" & NumEmpresa & "' "
  SelectAdodc AdoQuery, sSQL
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     PorcentajeProceso NombreTabla, Cont1
     LeerCamposTabla Cod_FieldTabla, Ninguno
     SetCamposTabla True
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarFecha(BCodigo As Variant, _
                           NumFile As Integer, _
                           NombreTabla As String)
Dim Cont1  As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  NumItemTemp = NumEmpresa
  Contador1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     Si_No = True
     PorcentajeProceso NombreTabla, Cont1
     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     sSQL = "SELECT * FROM " & NombreTabla & " " _
          & "WHERE " & BCodigo & " = '" & CodBusq & "' " _
          & "AND Fecha = #" & MiFecha & "# "
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
Dim Cont1  As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim SiEncontro As Boolean
  RatonReloj
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  NombreTabla = Trim(NombreTabla)
  Contador1 = 0
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
     PorcentajeProceso NombreTabla, Cont1
     LeerCamposTabla Cod_FieldTabla, Ninguno
     SetCamposTabla True
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarCodigo(BCodigo As Variant, _
                            NumFile As Integer, _
                            NombreTabla As String)
Dim Cont1  As Long
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
  Contador1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     PorcentajeProceso NombreTabla, Cont1
     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     With AdoQuery.Recordset
       If .RecordCount > 0 Then
          .MoveFirst
          .Find (BCodigo & " = '" & CodBusq & "' ")
       End If
       If Not .EOF Then
          SetCamposTabla False
       Else
        ' MsgBox .EOF
          SetCamposTabla True
       End If
     End With
     Cont1 = Cont1 + 1
  Loop
  RatonNormal
End Sub

Public Sub ActualizarCodigoN(BCodigo As Variant, _
                             NumFile As Integer, _
                             NombreTabla As String)
Dim Cont1  As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  sSQL = "SELECT * FROM " & NombreTabla & " " _
       & "WHERE " & BCodigo & " <> 0 " _
       & "ORDER BY " & BCodigo & " "
  SelectAdodc AdoQuery, sSQL
  Contador1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = 0
     PorcentajeProceso NombreTabla, Cont1
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
Dim Cont1  As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  NumItemTemp = NumEmpresa
  Contador1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     Si_No = True
     PorcentajeProceso NombreTabla, Cont1
     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     sSQL = "SELECT * " _
          & "FROM " & NombreTabla & " " _
          & "WHERE Credito_No = '" & Contrato_No & "' " _
          & "AND " & BCodigo & " = '" & CodBusq & "' " _
          & "AND Fecha = #" & MiFecha & "# " _
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
Dim Cont1  As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  NumItemTemp = NumEmpresa
  Contador1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     Si_No = True
     PorcentajeProceso NombreTabla, Cont1
     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     sSQL = "SELECT * " _
          & "FROM " & NombreTabla & " " _
          & "WHERE Cta = '" & Cta & "' " _
          & "AND " & BCodigo & " = '" & CodBusq & "' " _
          & "AND Item = '" & NumItemTemp & "' "
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

Public Sub ActualizarCodigoItem(BCodigo As Variant, _
                                NumFile As Integer, _
                                NombreTabla As String)
Dim Cont1  As Long
Dim Cod_FieldTabla As String
Dim Valor_Field As String
Dim CodBusq As Variant
Dim SiEncontro As Boolean
  RatonReloj
  BCodigo = Trim(BCodigo)
  NombreTabla = Trim(NombreTabla)
  NumItemTemp = NumEmpresa
  Contador1 = 0
  Do While Not EOF(NumFile)
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     Si_No = True
     PorcentajeProceso NombreTabla, Cont1
     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     sSQL = "SELECT * FROM " & NombreTabla & " " _
          & "WHERE " & BCodigo & " = '" & CodBusq & "' " _
          & "AND Item = '" & NumItemTemp & "' "
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

Public Sub ActualizarCodigoC(BCodigo As Variant, _
                             NumFile As Integer, _
                             NombreTabla As String)
Dim Cont1 As Long
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
  Contador1 = 0
  Do While Not EOF(NumFile)
     SiExist = True
     Line Input #NumFile, Cod_FieldTabla
     CodBusq = Ninguno
     NumItemTemp = NumEmpresa
     CodBusq = LeerCamposTabla(Cod_FieldTabla, BCodigo)
     PorcentajeProceso NombreTabla, Cont1
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

Private Sub CheqBCD_Change()
  If CheqBCD.Value = True Then CheqBT.Value = False
End Sub

Private Sub CheqBCE_Change()
  If CheqBCE.Value = True Then CheqBT.Value = False
End Sub

Private Sub CheqBCI_Change()
  If CheqBCI.Value = True Then CheqBT.Value = False
End Sub

Private Sub CheqBFA_Change()
  If CheqBFA.Value = True Then CheqBT.Value = False
End Sub

Private Sub CheqBNV_Change()
  If CheqBNV.Value = True Then CheqBT.Value = False
End Sub

Private Sub CheqBT_Change()
  If CheqBT.Value = True Then
     CheqBCD.Value = False
     CheqBCI.Value = False
     CheqBCE.Value = False
     CheqBFA.Value = False
     CheqBNV.Value = False
  End If
End Sub

Private Sub Command1_Click()
  ConSubDir = False
  If NumEmpresa = "" Then NumEmpresa = LGrupo.Text
  GrupoEmpresa = LGrupo.Text
  Contador = 0: FileResp = 0
  FechaValida MBFechaI
  FechaValida MBFechaF
' Eliminamos archivos de otros dias
  OpcRecibir.Value = True
  Dir1.Path = RutaBackup & "\DATOS\R" & LGrupo.Text
  Dir1.Refresh
  File1.FileName = Dir1.Path & "\F*.TXT"
  File1.Refresh
  If File1.ListCount > 0 Then Kill Dir1.Path & "\F*.TXT"
  RatonReloj
  FUnidad.Show 1
' MsgBox "Preceso Terminado"
  If RutaDestino <> "" Then
     Codigo = RutaDestino
     If (Codigo = "A:") Or (Codigo = "B:") Then
        RutaOrigen = Codigo & "\Z" _
                   & Format(Day(MBFechaI.Text), "00") _
                   & Format(Month(MBFechaI.Text), "00") _
                   & GrupoEmpresa & ".ZIP"
        RutaDestino = Dir1.Path & "\"
     Else
        RutaOrigen = Dir1.Path & "\Z" _
                   & Format(Day(MBFechaI.Text), "00") _
                   & Format(Month(MBFechaI.Text), "00") _
                   & GrupoEmpresa & ".ZIP"
        RutaDestino = Dir1.Path & "\"
     End If
     Cadena = Dir(RutaOrigen, vbArchive)
     If Cadena = "" Then
        MsgBox "Error: " & vbCrLf _
               & Space(9) & "El Archivo: " & vbCrLf _
               & Space(9) & RutaOrigen & vbCrLf _
               & Space(9) & "No Existe"
     Else
        Cadena = "SQLRESTA.BAT " _
               & Codigo & " " _
               & RutaOrigen & " " _
               & RutaDestino
        Shell Cadena, vbMaximizedFocus
     End If
  End If
  MsgBox "Fin de la descompresión"
  Dir1.Path = RutaBackup & "\DATOS\R" & LGrupo.Text
  Dir1.Refresh
  File1.FileName = Dir1.Path & "\F*.TXT"
  File1.Refresh
  RatonNormal
  Respaldos.Caption = "MODULO DE RESPALDOS"
  OpcRecibir.SetFocus
End Sub

Private Sub Command2_Click()
  Unload Respaldos
End Sub


'''Private Sub Command3_Click()
'''Dim Nombre1 As String
'''Dim Nombre2 As String
'''Dim Apellido1 As String
'''Dim Apellido2 As String
'''  EsAccess97 = False
'''  Titulo = "ELIMINACION"
'''  Mensajes = "En Access 97"
'''  If BoxMensaje = vbYes Then EsAccess97 = True
'''  If EsAccess97 Then
'''     Si_No = SQL_Server
'''     AdoStrCnnOld = AdoStrCnn
'''     SQL_Server = False
'''   ' Buscamos la cadena de conección a la base
'''     RutaGeneraFile = RutaSistema & "\CONECTAR.TXT"
'''     NumFile = FreeFile
'''     AdoStrCnn = ""
'''     Open RutaGeneraFile For Input As #NumFile
'''       Do While Not EOF(NumFile)
'''          AdoStrCnn = AdoStrCnn & Input(1, #NumFile) ' Obtiene un carácter.
'''       Loop
'''     Close #NumFile
'''     XAdoStrCnn = AdoStrCnn
'''     RutaEmpresa = UCase(RutaSistema & "\EMPRESA\" & Carpeta)
'''     RutaEmpresaOld = UCase(RutaSistema & "\EMPRESA\" & Carpeta)
'''   ' Procesamos las rutina necesarias antes de respaldar
'''     PathEmpresa = UCase(RutaSistema & "\EMPRESAS.MDB")
'''     AdoStrCnn = XAdoStrCnn & "Data Source=" & PathEmpresa
'''     ConectarAdodc AdoQuery
'''   ' Generamos Tabla:
'''     sSQL = "SELECT * " _
'''          & "FROM Empresas " _
'''          & "WHERE Item <> 0 " _
'''          & "ORDER BY Item "
'''     SelectData AdoQuery, sSQL
'''     If AdoQuery.Recordset.RecordCount > 0 Then
'''        AdoQuery.Recordset.MoveFirst
'''        AdoQuery.Recordset.Find ("Item = " & Val(NumEmpresa) & " ")
'''        If Not AdoQuery.Recordset.EOF Then
'''           Carpeta = AdoQuery.Recordset.Fields("SubDir")
'''           PathEmpresa = UCase(RutaSistema & "\EMPRESA\" & Carpeta)
'''        End If
'''     End If
'''  End If
'''' Procesamos las rutina necesarias antes de respaldar
'''  If EsAccess97 Then
'''     AdoStrCnn = XAdoStrCnn & "Data Source=" & PathEmpresa & "\ENVIOS.MDB"
'''     ConectarAdodc AdoAux
'''     ConectarAdodc AdoAct
'''     ConectarAdodc AdoQuery
'''  End If
'''  If EsAccess97 Then
'''     EliminarRepetidosDe True, "Beneficiarios", "Codigo_B"
'''  Else
'''     EliminarRepetidosDe False, "Clientes", "Codigo"
'''  End If
'''  If EsAccess97 Then
'''     EliminarRepetidosDe True, "Remitentes", "Codigo_R"
'''  Else
'''     EliminarRepetidosDe False, "Remitentes", "Codigo_R"
'''  End If
'''  EliminarRepetidosDe False, "Correos", "Envio_No"
'''' Restauramos datos
'''  If EsAccess97 Then
'''     AdoStrCnn = AdoStrCnnOld
'''     SQL_Server = Si_No
'''  End If
'''  RatonNormal
'''  MsgBox "Proceso Terminado"
'''End Sub

Private Sub Command4_Click()
  If NumEmpresa = "" Then NumEmpresa = LGrupo.Text
  ConSubDir = False
  If RutaDestino <> "" Then TipoRespaldo False
  Unload Respaldos
End Sub


Private Sub Command5_Click()
Dim AuxNumEmp As String
  AuxNumEmp = NumEmpresa
  NumEmpresa = Trim(Cod_NumEmp)
  If NumEmpresa = "" Then NumEmpresa = LGrupo.Text
  OpcRecibir.Value = True
  Dir1.Path = RutaBackup & "\DATOS\R" & LGrupo.Text
  Dir1.Refresh
  File1.FileName = Dir1.Path & "\F*.TXT"
  File1.Refresh
  Contador = 0: FileResp = 0
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  ConectarAdodc AdoAct
  ConectarAdodc AdoAux
  ConectarAdodc AdoQuery
  'SelectAdodc AdoQuery, sSQL
  TextArchivo.Text = ""
  ProgBarra.Value = 0
  ProgBarra.Max = File1.ListCount + 10
  For IJ = 0 To File1.ListCount - 1
      RatonReloj
      ProgBarra.Value = ProgBarra.Value + 1
      TextArchivo.Text = TextArchivo.Text & Space(9) & File1.List(IJ) & " => "
      RutaGeneraFile = RutaSysBases & "\DATOS\R" & GrupoEmpresa & "\" & File1.List(IJ)
      ' MsgBox RutaSysBases
      NumFile = FreeFile
      Open RutaGeneraFile For Input As #NumFile
           AbrirCamposSQL NumFile
           'Respaldos.Caption = Cod_Base & ": Procesando(" & 1 & ") " & String(1, "|")
           TextArchivo.Text = TextArchivo.Text & Cod_Base & vbCrLf
           TextArchivo.Refresh
           Respaldos.Refresh
           Select Case Cod_Base
             Case "Fecha_Respaldo"
                  AbrirArchivoSQL NumFile
                  TextArchivo.Text = TextArchivo.Text _
                                   & "PROCESOS REALIZADOS:" & vbCrLf _
                                   & "===================" & vbCrLf
                  TextArchivo.Refresh
             Case "Accesos": ActualizarCodigoItem "Codigo", NumFile, Cod_Base
                             'ActualizarCodigo "Codigo", NumFile
             Case "Acceso_Empresa": ActualizarCodigo "Codigo", NumFile, Cod_Base
             Case "Codigos": ActualizarCodigoItem "Concepto", NumFile, "Codigos"
             Case "Ctas_Proceso": If ConSucursal = False Then ActualizarCodigoItem "Detalle", NumFile, Cod_Base
             Case "Catalogo_Cuentas": If ConSucursal = False Then ActualizarCodigoItem "Codigo", NumFile, Cod_Base
             Case "Catalogo_SubCtas": ActualizarCodigoItem "Codigo", NumFile, Cod_Base
             Case "Catalogo_CxCxP": ActualizarCodigoCta "Codigo", NumFile, Cod_Base
             Case "Catalogo_RolPagos": ActualizarCodigoItem "Codigo", NumFile, Cod_Base
             Case "Catalogo_Productos": ActualizarCodigoItem "Codigo_Inv", NumFile, Cod_Base
             Case "Comprobantes": ActualizarRangoFecha NumFile, Cod_Base
             Case "Transacciones": ActualizarRangoFecha NumFile, Cod_Base
             Case "Trans_Abonos": ActualizarRangoFecha NumFile, Cod_Base
             Case "Trans_Bancos": ActualizarRangoFecha NumFile, Cod_Base
             Case "Trans_SubCtas": ActualizarRangoFecha NumFile, Cod_Base
             Case "Trans_Kardex": ActualizarRangoFecha NumFile, Cod_Base
             Case "Trans_Retenciones": ActualizarRangoFecha NumFile, Cod_Base
             Case "Trans_Dep_Chq": ActualizarRangoFecha NumFile, Cod_Base
             Case "Trans_Libretas": ActualizarRangoFecha NumFile, Cod_Base
             Case "Trans_Prestamos": ActualizarMayor NumFile, "Trans_Prestamos"
             Case "Trans_PrestamosC": ActualizarCodigoP "Cuenta_No", NumFile, "Trans_Prestamos"
             Case "Trans_Cajas": ActualizarRangoFecha NumFile, Cod_Base
             Case "Bloqueos": ActualizarRangoFecha NumFile, Cod_Base
             Case "Prestamos": ActualizarMayor NumFile, Cod_Base
             Case "Cuentas": ActualizarRangoFecha NumFile, Cod_Base
             Case "Saldo_Caja_Libreta": ActualizarRangoFecha NumFile, Cod_Base
             Case "Saldo_Libretas_Intereses": ActualizarRangoFecha NumFile, Cod_Base
             Case "Correos": ActualizarCodigoC "Envio_No", NumFile, "Correos"
             Case "Clientes": ActualizarCodigo "Codigo", NumFile, "Clientes"
             Case "Remitentes": ActualizarCodigo "Codigo_R", NumFile, "Remitentes"
             Case "Garantes": ActualizarMayor NumFile, Cod_Base
             Case "Resumen_Llamadas": ActualizarRangoFecha NumFile, "Resumen_Llamadas"
             Case "Abono_De_Prestamo": ActualizarCodigoItem "Cuenta", NumFile, "Abono_De_Prestamo"
             Case "Asiento_De_Prestamo": ActualizarCodigoItem "Cuenta", NumFile, "Asiento_De_Prestamo"
             Case "Tipo_Prestamo": ActualizarCodigoItem "TP", NumFile, "Tipo_Prestamo"
             Case "Conyugue": ActualizarRangoFecha NumFile, Cod_Base
             Case "Detalle_Factura": ActualizarRangoFecha NumFile, Cod_Base
             Case "Facturas": ActualizarRangoFecha NumFile, Cod_Base
             Case "Prestamos": ActualizarRangoFecha NumFile, Cod_Base
           End Select
      Close #NumFile
      RatonNormal
  Next IJ
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  Respaldos.Caption = "MODULO DE RESPALDOS"
  NumEmpresa = AuxNumEmp
  MsgBox "Fin del Proceso"
End Sub

Private Sub Command6_Click()
  If NumEmpresa = "" Then NumEmpresa = LGrupo.Text
  'FUnidad.Show 1
  ConSubDir = False
  If RutaDestino <> "" Then TipoRespaldo True
  Unload Respaldos
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
  'If KeyCode = vbKeyReturn Then EjecutarRespaldo
  If KeyCode = vbKeyDelete Then
     Mensajes = "Esta seguro de Eliminar: " & File1.FileName
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then Kill File1.Path & "\" & File1.FileName
     File1.FileName = Dir1.Path & "\*.TXT"
  End If
End Sub

Private Sub Form_Activate()
  FechaValida MBFechaI
  FechaValida MBFechaF
  CheqBT.Value = True
' Generamos Tabla:
  LGrupo.Clear
  sSQL = "SELECT Grupo As Grupo_No " _
       & "FROM Empresas " _
       & "WHERE Grupo <> '000' " _
       & "GROUP BY Grupo "
  SelectData AdoQuery, sSQL
  RatonReloj
  With AdoQuery.Recordset
    Do While Not .EOF
       LGrupo.AddItem .Fields("Grupo_No")
       Codigo = RutaSysBases & "\DATOS\E" & .Fields("Grupo_No")
       Codigo1 = RutaSysBases & "\DATOS\R" & .Fields("Grupo_No")
       Cadena = Dir(Codigo, vbDirectory)
       If Cadena = "" Then MkDir (Codigo)
       Cadena = Dir(Codigo1, vbDirectory)
       If Cadena = "" Then MkDir (Codigo1)
      .MoveNext
    Loop
  End With
  LGrupo.Text = NumEmpresa
  Drive1.Drive = Mid(RutaSysBases, 1, 2)
  RatonNormal
  Codigo = Mid(MBFechaI.Text, 1, 2) _
         & Mid(MBFechaI.Text, 4, 2)
  RutaBackup = RutaSysBases
  Dir1.Path = RutaBackup & "\DATOS\E" & GrupoEmpresa
  File1.FileName = Dir1.Path & "\F*.TXT"
  TextArchivo.Text = NumEmpresa & ", Carpeta Base Anterior: " & Carpeta
  Respaldos.Caption = "MODULO DE RESPALDOS"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm Respaldos
  Command6.Enabled = False
  If CodigoUsuario = "ACCESO02" Then Command6.Enabled = True
  ConectarAdodc AdoAux
  ConectarAdodc AdoAct
  ConectarAdodc AdoQuery
End Sub

Private Sub LGrupo_LostFocus()
  GrupoEmpresa = LGrupo.Text
  If OpcEnviar.Value Then
     Dir1.Path = RutaBackup & "\DATOS\E" & LGrupo.Text
  Else
     Dir1.Path = RutaBackup & "\DATOS\R" & LGrupo.Text
  End If
  Dir1.Refresh
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If CtrlDown And KeyCode = vbKeyF10 Then
     TextArchivo.Text = ""
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
     TextArchivo.Refresh
  End If
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
End Sub

Private Sub MBFechaI_GotFocus()
  'MsgBox "|" & CompilarString(InputBox("Cadena:", "Texto")) & "|"
  TextArchivo.Text = ""
  If File1.ListCount > 0 Then
     RatonReloj
     If OpcRecibir.Value Then
        Cod_FechaI = "": Cod_FechaF = ""
        RutaGeneraFile = RutaSysBases & "\DATOS\R" & GrupoEmpresa & "\" & File1.List(0)
        Cadena = Dir(RutaGeneraFile, vbArchive)
        If Cadena <> "" Then
           NumFile = FreeFile
           Open RutaGeneraFile For Input As #NumFile
                Line Input #NumFile, Cadena
                Line Input #NumFile, Cadena
                AbrirArchivoSQL NumFile
           Close #NumFile
        End If
     End If
     RatonNormal
  End If
  TextArchivo.Refresh
  If Cod_FechaI = "" Then Cod_FechaI = FechaSistema
  If Cod_FechaF = "" Then Cod_FechaF = FechaSistema
  MBFechaI.Text = Cod_FechaI
  MBFechaF.Text = Cod_FechaF
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  File1.FileName = Dir1.Path & "\F*.TXT"
  File1.Refresh
  TextArchivo.Refresh
End Sub

Private Sub OpcEnviar_Click()
  If OpcEnviar.Value Then
     Dir1.Path = RutaBackup & "\DATOS\E" & LGrupo.Text
  Else
     Dir1.Path = RutaBackup & "\DATOS\R" & LGrupo.Text
  End If
  Dir1.Refresh
End Sub

Private Sub OpcRecibir_Click()
  If OpcRecibir.Value Then
     Dir1.Path = RutaBackup & "\DATOS\R" & LGrupo.Text
  Else
     Dir1.Path = RutaBackup & "\DATOS\E" & LGrupo.Text
  End If
  Dir1.Refresh
End Sub

Public Function LeerCamposTabla(Cod_FieldT1 As String, _
                                BCodigo As Variant) As Variant
Dim CodBusp1 As Variant
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
      If TipoC(I).Campo = "Fecha" Then MiFecha = BuscarFecha(Format(Valor_Field, FormatoFechas))
      If TipoC(I).Campo = "Mes_No" Then NoMeses = Valor_Field
      If TipoC(I).Campo = "TP" Then TipoProc = Valor_Field
      If TipoC(I).Campo = "Credito_No" Then Contrato_No = Valor_Field
      If TipoC(I).Campo = "Item" Then NumItemTemp = Valor_Field
      If TipoC(I).Campo = "Cta" Then Cta = TipoC(I).Valor
      If TipoC(I).Campo = BCodigo Then CodBusq1 = TipoC(I).Valor
   End If
   Cod_FieldT1 = Mid(Cod_FieldT1, No_Hasta + 1, Len(Cod_FieldT1))
   No_Desde = 1: No_Hasta = 1
  Next I
  LeerCamposTabla = CodBusq1
End Function

Public Sub SetCamposTabla(FAddNew As Boolean)
  With AdoQuery.Recordset
   If FAddNew Then SetAddNew AdoQuery
   For J = 0 To .Fields.Count - 1
       SiEncontro = False: I = 1
       Do
         If .Fields(J).Name = TipoC(I).Campo Then
             SetFields AdoQuery, TipoC(I).Campo, TipoC(I).Valor
             SiEncontro = True
         End If
         I = I + 1
       Loop Until I > CantCampos
       If SiEncontro = False Then
          Select Case .Fields(J).Type
            Case TadBoolean
                 SetFields AdoQuery, .Fields(J).Name, False
            Case TadDate, TadDate1
                 SetFields AdoQuery, .Fields(J).Name, FechaSistema
            Case TadTime
                 SetFields AdoQuery, .Fields(J).Name, TiempoSistema
            Case TadByte, TadInteger, TadLong, TadDouble, TadSingle, TadCurrency
                 SetFields AdoQuery, .Fields(J).Name, 0
            Case TadText
                 SetFields AdoQuery, .Fields(J).Name, Ninguno
            Case Else
                 SetFields AdoQuery, .Fields(J).Name, Ninguno
          End Select
       End If
   Next J
   SetUpdate AdoQuery
  End With
End Sub

Public Sub PorcentajeProceso(NombTabla As String, _
                             ContX As Long)
  Cadena = NombTabla _
         & ": Procesando(" _
         & Format(ContX / TotalReg, "##0%") _
         & ") " _
         & String(ContX Mod 40, "|")
  Respaldos.Caption = Cadena
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

Public Sub PrepararClientes97()
  sSQL = "DELETE * " _
       & "FROM Clientes " _
       & "WHERE Grupo <> '999999' "
  ConectarAdoExecute sSQL
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Grupo <> '.' "
  SelectAdodc AdoAux, sSQL
  Contador = 0
  Respaldos.Caption = "Procesando Codigo de Clientes en Comprobantes..."
 'Beneficiarios
  sSQL = "SELECT * " _
       & "FROM Beneficiarios " _
       & "WHERE TC <> 'X' " _
       & "ORDER BY Beneficiario "
  SelectAdodc AdoAct, sSQL
  With AdoAct.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          Respaldos.Caption = "Procesando " & Format(Contador / .RecordCount, "00%") & ", Codigo de Clientes en Beneficiarios..."
          SetFields AdoAct, "Beneficiario", UCase(CompilarString(.Fields("Beneficiario")))
          SetFields AdoAct, "RUC_CI", CompilarRUC_CI(.Fields("RUC_CI"))
          SetFields AdoAct, "Direccion", UCase(CompilarString(.Fields("Direccion")))
          SetFields AdoAct, "Telefono", UCase(CompilarString(.Fields("Telefono")))
          SetFields AdoAct, "Celular", UCase(CompilarString(.Fields("Celular")))
          SetFields AdoAct, "FAX", UCase(CompilarString(.Fields("FAX")))
          SetFields AdoAct, "Ciudad", UCase(CompilarString(.Fields("Ciudad")))
          SetUpdate AdoAct
         .MoveNext
       Loop
   End If
  End With
  Respaldos.Caption = "Procesando Codigo de Clientes en Facturacion..."
 'Clientes en Facturacion
  sSQL = "SELECT * " _
       & "FROM Clientes1 " _
       & "WHERE E <> 'X' " _
       & "ORDER BY Cliente "
  SelectAdodc AdoAct, sSQL
  With AdoAct.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          Respaldos.Caption = "Procesando " & Format(Contador / .RecordCount, "00%") & ", Codigo de Clientes en Facturacion..."
          SetFields AdoAct, "Cliente", UCase(CompilarString(.Fields("Cliente")))
          SetFields AdoAct, "RUC_CI", CompilarRUC_CI(.Fields("RUC_CI"))
          SetFields AdoAct, "Direccion", UCase(CompilarString(.Fields("Direccion")))
          SetFields AdoAct, "Telefono", UCase(CompilarString(.Fields("Telefono")))
          SetFields AdoAct, "Celular", UCase(CompilarString(.Fields("Celular")))
          SetFields AdoAct, "FAX", UCase(CompilarString(.Fields("FAX")))
          SetFields AdoAct, "Ciudad", UCase(CompilarString(.Fields("Ciudad")))
          SetUpdate AdoAct
         .MoveNext
       Loop
   End If
  End With
 'Comprobantes
  Contador = 0
  sSQL = "SELECT * " _
       & "FROM Comprobantes " _
       & "WHERE T <> 'X' " _
       & "ORDER BY Beneficiario "
  SelectAdodc AdoAct, sSQL
  With AdoAct.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          Respaldos.Caption = "Procesando " & Format(Contador / .RecordCount, "00%") & ", Codigo de Clientes en Comprobantes..."
          SetFields AdoAct, "Beneficiario", UCase(CompilarString(.Fields("Beneficiario")))
          SetFields AdoAct, "RUC_CI", CompilarRUC_CI(.Fields("RUC_CI"))
          SetUpdate AdoAct
         .MoveNext
       Loop
   End If
  End With
 'Comprobantes
  Contador = 0
  sSQL = "SELECT Beneficiario " _
       & "FROM Comprobantes " _
       & "WHERE T <> 'X' " _
       & "GROUP BY Beneficiario " _
       & "ORDER BY Beneficiario "
  SelectAdodc AdoAct, sSQL
  With AdoAct.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          Respaldos.Caption = "Insertando Comprobantes " & Format(Contador / .RecordCount, "00%") & "..."
          SetAddNew AdoAux
          SetFields AdoAux, "T", "N"
          SetFields AdoAux, "Codigo", NumEmpresa & Format(Contador, "0000000")
          SetFields AdoAux, "Cliente", .Fields("Beneficiario")
          SetFields AdoAux, "CI_RUC", "9000000000"
          SetFields AdoAux, "Grupo", NumEmpresa
          SetFields AdoAux, "TD", "O"
          SetUpdate AdoAux
         .MoveNext
       Loop
   End If
  End With
  sSQL = "UPDATE Comprobantes As Co,Clientes As C " _
       & "SET Co.CodigoB = C.Codigo " _
       & "WHERE Co.Beneficiario = C.Cliente "
  ConectarAdoExecute sSQL

  sSQL = "UPDATE Clientes As C,Comprobantes As Co " _
       & "SET C.Fecha = Co.Fecha," _
       & "C.CI_RUC = Co.RUC_CI " _
       & "WHERE C.Codigo = Co.CodigoB "
  ConectarAdoExecute sSQL
 'Beneficiarios
  sSQL = "SELECT * " _
       & "FROM Beneficiarios " _
       & "WHERE TC <> 'X' " _
       & "ORDER BY Beneficiario "
  SelectAdodc AdoAct, sSQL
  With AdoAct.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          Respaldos.Caption = "Insertando Beneficiarios " & Format(Contador / .RecordCount, "00%") & "..."
          SetAddNew AdoAux
          SetFields AdoAux, "T", "N"
          SetFields AdoAux, "Codigo", String(2, .Fields("TC")) & .Fields("Codigo")
          SetFields AdoAux, "Cliente", .Fields("Beneficiario")
          SetFields AdoAux, "CI_RUC", .Fields("RUC_CI")
          SetFields AdoAux, "Direccion", .Fields("Direccion")
          SetFields AdoAux, "Ciudad", .Fields("Ciudad")
          SetFields AdoAux, "Telefono", .Fields("Telefono")
          SetFields AdoAux, "Celular", .Fields("Celular")
          SetFields AdoAux, "FAX", .Fields("FAX")
          SetFields AdoAux, "Grupo", NumEmpresa
          SetFields AdoAux, "TD", "O"
          SetUpdate AdoAux
         .MoveNext
       Loop
   End If
  End With
 'Clientes en Facturacion
  Contador = 0
  sSQL = "SELECT * " _
       & "FROM Clientes1 " _
       & "WHERE E <> 'X' " _
       & "ORDER BY Cliente "
  SelectAdodc AdoAct, sSQL
  With AdoAct.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Contador = Contador + 1
          Respaldos.Caption = "Insertando Clientes en Facturacion " & Format(Contador / .RecordCount, "00%") & "..."
          SetAddNew AdoAux
          SetFields AdoAux, "T", "N"
          SetFields AdoAux, "Codigo", "FA" & .Fields("Codigo")
          SetFields AdoAux, "Cliente", .Fields("Cliente")
          SetFields AdoAux, "CI_RUC", .Fields("RUC_CI")
          SetFields AdoAux, "Direccion", .Fields("Direccion")
          SetFields AdoAux, "Ciudad", .Fields("Ciudad")
          SetFields AdoAux, "Telefono", .Fields("Telefono")
          SetFields AdoAux, "Celular", .Fields("Celular")
          SetFields AdoAux, "FAX", .Fields("FAX")
          SetFields AdoAux, "Fecha_N", .Fields("Fecha_N")
          SetFields AdoAux, "Profesion", .Fields("Profesion")
          SetFields AdoAux, "Representante", .Fields("Empresa")
          SetFields AdoAux, "Email", .Fields("Email")
          SetFields AdoAux, "Sexo", .Fields("Sexo")
          SetFields AdoAux, "Grupo", .Fields("Grupo")
          SetFields AdoAux, "TD", "O"
          SetUpdate AdoAux
         .MoveNext
       Loop
   End If
  End With
  Respaldos.Caption = "Actualizando Campos de Clientes..."
  sSQL = "UPDATE Clientes SET Direccion = 'S/N' WHERE Direccion = '.' "
  ConectarAdoExecute sSQL
  sSQL = "UPDATE Clientes SET DirNumero = 'S/N' WHERE DirNumero = '.' "
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
