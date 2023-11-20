VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AcEntSal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizando Entradas/Salidas"
   ClientHeight    =   360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   6315
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   2730
      TabIndex        =   3
      Top             =   0
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
      Height          =   330
      Left            =   5145
      Picture         =   "AcEntSal.frx":0000
      TabIndex        =   5
      Top             =   0
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Procesar"
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
      Left            =   3990
      Picture         =   "AcEntSal.frx":08CA
      TabIndex        =   4
      Top             =   0
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   1575
      Top             =   420
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSAdodcLib.Adodc AdoAux1 
      Height          =   330
      Left            =   3255
      Top             =   420
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Aux1"
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
      Left            =   735
      TabIndex        =   1
      Top             =   0
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
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Hasta"
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
      Left            =   1995
      TabIndex        =   2
      Top             =   0
      Width           =   750
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Desde"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   750
   End
End
Attribute VB_Name = "AcEntSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ActualizarEntradasSalidas()
  SQL2 = "DELETE * " _
       & "FROM Trans_RolHoras " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  ConectarAdoExecute SQL2
  
  SQL2 = "SELECT * " _
       & "FROM Trans_RolHoras " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectAdodc AdoAux, SQL2
  
  sSQL = "SELECT Codigo,Fecha,Hora,Count(Codigo) As NReg " _
       & "FROM Trans_Entrada_Salida " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND ES = 'H' " _
       & "GROUP BY Codigo,Fecha,Hora " _
       & "ORDER BY Codigo,Fecha,Hora "
  SelectAdodc AdoAux1, sSQL
  With AdoAux1.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       Codigo = .Fields("Codigo")
       MiFecha = .Fields("Fecha")
       HoraDiaI = CDate(.Fields("Hora"))
       Do While Not .EOF
          Contador = Contador + 1
          AcEntSal.Caption = "Actualizando Entradas/Salidas: " & Format(Contador / .RecordCount, "00%")
          If Codigo <> .Fields("Codigo") Or MiFecha <> .Fields("Fecha") Then
             Real1 = Hour(HoraDiaF - HoraDiaI) & "." & Minute(HoraDiaF - HoraDiaI)
             If Real1 <= 0 Then Real1 = 0.01
             SetAddNew AdoAux
             SetFields AdoAux, "Codigo", Codigo
             SetFields AdoAux, "Fecha", MiFecha
             SetFields AdoAux, "Horas", Real1
             SetFields AdoAux, "Item", NumEmpresa
             SetUpdate AdoAux
             Codigo = .Fields("Codigo")
             MiFecha = .Fields("Fecha")
             HoraDiaI = CDate(.Fields("Hora"))
          End If
          HoraDiaF = CDate(.Fields("Hora"))
         .MoveNext
       Loop
       Real1 = Hour(HoraDiaF - HoraDiaI) & "." & Minute(HoraDiaF - HoraDiaI)
       If Real1 <= 0 Then Real1 = 0.01
       SetAddNew AdoAux
       SetFields AdoAux, "Codigo", Codigo
       SetFields AdoAux, "Fecha", MiFecha
       SetFields AdoAux, "Horas", Real1
       SetFields AdoAux, "Item", NumEmpresa
       SetUpdate AdoAux
   End If
  End With
End Sub

Public Sub ActualizarVentasAnticipadas()

End Sub

Private Sub Command1_Click()
Dim HoraDiaI As Date
Dim HoraDiaF As Date
 'Consultar
  RatonReloj
  Contador = 0
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  Select Case Modulo
    Case "FACTURAS": ActualizarVentasAnticipadas
    Case "ROL PAGOS": ActualizarEntradasSalidas
  End Select
  RatonNormal
  Unload AcEntSal
End Sub

Private Sub Command3_Click()
  Unload AcEntSal
End Sub

Private Sub Form_Load()
  Select Case Modulo
    Case "FACTURAS": AcEntSal.Caption = "ACTUALIZANDO VENTAS ANTICIPADAS"
    Case "ROL PAGOS": AcEntSal.Caption = "ACTUALIZANDO ENTRADAS/SALIDAS"
  End Select
  CentrarForm AcEntSal
  ConectarAdodc AdoAux
  ConectarAdodc AdoAux1
  RatonNormal
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

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
End Sub

