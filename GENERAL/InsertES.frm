VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form InsertarEntSal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Comisiones y el I.E.S.S."
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   Icon            =   "InsertES.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox MBEntrada 
      Height          =   330
      Left            =   1470
      TabIndex        =   3
      Top             =   525
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "HH:MM"
      Mask            =   "##:##"
      PromptChar      =   "0"
   End
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   105
      Top             =   1260
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
      Caption         =   "CxCxP"
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
      Left            =   2520
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
   Begin VB.CheckBox CheqDom 
      Caption         =   "Domingos"
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
      Left            =   1575
      TabIndex        =   7
      Top             =   945
      Width           =   1275
   End
   Begin VB.CheckBox CheqSab 
      Caption         =   "Sábados"
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
      TabIndex        =   6
      Top             =   945
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Procesar Mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4515
      Picture         =   "InsertES.frx":0696
      TabIndex        =   8
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
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
      Height          =   540
      Left            =   4515
      Picture         =   "InsertES.frx":0AD8
      TabIndex        =   9
      Top             =   735
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoDias 
      Height          =   330
      Left            =   1995
      Top             =   1260
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
      Caption         =   "Dias"
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
   Begin MSMask.MaskEdBox MBHora 
      Height          =   330
      Left            =   3990
      TabIndex        =   5
      Top             =   525
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "0"
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Intervalos de (ss)"
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
      Left            =   2310
      TabIndex        =   4
      Top             =   525
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Hora Entrada"
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
      Top             =   525
      Width           =   1380
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Primer día mes a procesar"
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
      Width           =   2430
   End
End
Attribute VB_Name = "InsertarEntSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub IngresoLabores(TextoTarea As String)
  SetAdoAddNew "Trans_Entrada_Salida"
  SetAdoFields "ES", "H"
  SetAdoFields "Codigo", CodigoCli
  SetAdoFields "Hora", TiempoTexto
  SetAdoFields "Fecha", FechaInicial
  SetAdoFields "Tarea", TextoTarea
  SetAdoFields "CodigoU", CodigoUsuario
  SetAdoFields "Item", NumEmpresa
  SetAdoUpdate
End Sub

Private Sub Command1_Click()
  FechaValida MBFechaI
  FechaFinal = UltimoDiaMes(MBFechaI.Text)
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(FechaFinal)
  'MsgBox FechaFin
  sSQL = "DELETE * " _
       & "FROM Trans_Entrada_Salida " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# AND #" & FechaFin & "# " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  Ejecutar_SQL_SP sSQL
  OpcTM = 5
  Contador = 0
  If CheqDom.value = 1 Then OpcTM = OpcTM + 1
  If CheqSab.value = 1 Then OpcTM = OpcTM + 1
  With AdoCxCxP.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       FechaInicial = MBFechaI.Text
       IE = 0
       SabadosMes = 0: DomingosMes = 0
       For I = 1 To DatePart("d", FechaFinal)        ' Ultimo dia mes
           Select Case DatePart("w", FechaInicial)   ' Dia de la semana
             Case 1: DomingosMes = DomingosMes + 1
             Case 7: SabadosMes = SabadosMes + 1
             Case Else: IE = IE + 1
           End Select
           FechaInicial = CLongFecha(CFechaLong(FechaInicial) + 1)
       Next I
       JE = IE
       If CheqDom.value = 1 Then JE = JE + DomingosMes
       If CheqSab.value = 1 Then JE = JE + SabadosMes
       Do While Not .EOF
          InsertarEntSal.Caption = .RecordCount & ", ASIGNAR DIAS LABORADOS: " & Format(Contador / .RecordCount, "00%")
          Contador = Contador + 1
          CodigoCli = .Fields("Codigo")
          Salida = Round(.Fields("Horas_Sem") / OpcTM, 2)
          Salida = (Salida * Val(MBHora.Text)) / 60
         'MsgBox SumaHora(MBEntrada.Text, 0) & vbCrLf & SumaHora(MBEntrada, Salida) & vbCrLf & Salida
          If JE > 0 Then
             FechaInicial = MBFechaI.Text
             For I = 1 To DatePart("d", FechaFinal)
                 TiempoTexto = Format(MBEntrada.Text, FormatoTimes)
                 Select Case DatePart("w", FechaInicial)
                   Case 1: If CheqDom.value = 1 Then IngresoLabores "Entrada"
                   Case 7: If CheqSab.value = 1 Then IngresoLabores "Entrada"
                   Case Else: IngresoLabores "Entrada"
                 End Select
                 TiempoTexto = SumaHora(MBEntrada.Text, CSng(Salida))
                 'MsgBox TiempoTexto
                 Select Case DatePart("w", FechaInicial)
                   Case 1: If CheqDom.value = 1 Then IngresoLabores "Salida"
                   Case 7: If CheqSab.value = 1 Then IngresoLabores "Salida"
                   Case Else: IngresoLabores "Salida"
                 End Select
                 FechaInicial = CLongFecha(CFechaLong(FechaInicial) + 1)
             Next I
          End If
         .MoveNext
       Loop
   End If
  End With
  Unload Me
End Sub

Private Sub Command3_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT * " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Usuario,Clave "
  Select_Adodc AdoCxCxP, sSQL
  MBHora.Text = "60"
  MBFechaI.SetFocus
End Sub

Private Sub Form_Load()
  CentrarForm InsertarEntSal
  ConectarAdodc AdoDias
  ConectarAdodc AdoCxCxP
  InsertarEntSal.Caption = "REGISTRO DE ENTRADA/SALIDA GLOBAL"
End Sub

Private Sub MBEntrada_GotFocus()
  MarcarTexto MBEntrada
End Sub

Private Sub MBEntrada_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
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

