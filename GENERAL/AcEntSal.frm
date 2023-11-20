VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form AcEntSal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizando Entradas/Salidas"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4830
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
      Left            =   2205
      TabIndex        =   7
      Top             =   210
      Width           =   1170
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
      Left            =   2205
      TabIndex        =   6
      Top             =   525
      Width           =   1275
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   420
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
      Left            =   3570
      Picture         =   "AcEntSal.frx":0000
      TabIndex        =   5
      Top             =   420
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
      Left            =   3570
      Picture         =   "AcEntSal.frx":08CA
      TabIndex        =   4
      Top             =   105
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   0
      Top             =   735
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
      Left            =   1785
      Top             =   735
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
      Left            =   840
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
      Left            =   105
      TabIndex        =   2
      Top             =   420
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
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   750
   End
End
Attribute VB_Name = "AcEntSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Actualizar_Entradas_Salidas()
Dim HoraDiaI As Date
Dim HoraDiaF As Date
  SQL2 = "DELETE * " _
       & "FROM Trans_Rol_Horas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' "
  Ejecutar_SQL_SP SQL2
  SQL2 = "SELECT * " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY Codigo "
  Select_Adodc AdoAux, SQL2
  sSQL = "SELECT Codigo,Fecha,Hora,Count(Codigo) As NReg " _
       & "FROM Trans_Entrada_Salida " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND ES = 'H' " _
       & "GROUP BY Codigo,Fecha,Hora " _
       & "ORDER BY Codigo,Fecha,Hora "
  Select_Adodc AdoAux1, sSQL
  With AdoAux1.Recordset
   If .RecordCount > 0 Then
       RatonReloj
       Codigo = .Fields("Codigo")
       Mifecha = .Fields("Fecha")
       HoraDiaI = CDate(.Fields("Hora"))
       HoraDiaF = CDate(.Fields("Hora"))
       Do While Not .EOF
          Contador = Contador + 1
          AcEntSal.Caption = "Actualizando Entradas/Salidas: " & Format$(Contador / .RecordCount, "00%")
          If Codigo <> .Fields("Codigo") Or Mifecha <> .Fields("Fecha") Then
             Real2 = 0
             If AdoAux.Recordset.RecordCount > 0 Then
                AdoAux.Recordset.MoveFirst
                AdoAux.Recordset.Find ("Codigo = '" & Codigo & "' ")
                If Not AdoAux.Recordset.EOF Then
                   Opciones = Val(AdoAux.Recordset.Fields("SN"))
                   Real2 = AdoAux.Recordset.Fields("Valor_Hora")
                   Cantidad = AdoAux.Recordset.Fields("Horas_Sem") / 5
                   Real1 = Val(Hour(HoraDiaF - HoraDiaI) & "." & Minute(HoraDiaF - HoraDiaI))
                   If Opciones = 2 Then
                      If Real1 < Cantidad Then Real1 = Cantidad
                      If Real1 > Cantidad Then Real1 = Cantidad
                   End If
                   If HoraDiaF = HoraDiaI Then Real1 = Cantidad
                   If Real1 <= 0 Then Real1 = 0.01
                   Real3 = Redondear(Real1 * Real2, 2)
                   SetAdoAddNew "Trans_Rol_Horas"
                   SetAdoFields "T", Normal
                   SetAdoFields "Codigo", Codigo
                   SetAdoFields "Fecha", Mifecha
                   SetAdoFields "Horas", Real1
                   SetAdoFields "Valor_Hora", Real2
                   SetAdoFields "Ing_Liquido", Real3
                   SetAdoFields "Fondo_Reserva", Redondear(Real3 / 12, 4)
                   SetAdoFields "IESS_Per", Redondear(Real3 * AdoAux.Recordset.Fields("IEESS_Per"), 4)
                   SetAdoFields "IESS_Pat", Redondear(Real3 * AdoAux.Recordset.Fields("IEESS_Pat"), 4)
                   SetAdoFields "Aporte_Adm", Redondear(Real3 * AdoAux.Recordset.Fields("Aporte_Adm"), 4)
                   SetAdoFields "Certificado", Redondear(Real3 * AdoAux.Recordset.Fields("Aporte_Cer"), 4)
                   SetAdoFields "Fondo_Emergencia", Redondear(Real3 * AdoAux.Recordset.Fields("Aporte_Fon"), 4)
                   SetAdoFields "Item", NumEmpresa
                   SetAdoUpdate
                   Codigo = .Fields("Codigo")
                   Mifecha = .Fields("Fecha")
                   HoraDiaI = CDate(.Fields("Hora"))
                   HoraDiaF = CDate(.Fields("Hora"))
                End If
             End If
          End If
          HoraDiaF = CDate(.Fields("Hora"))
         .MoveNext
       Loop
       Real2 = 0
             If AdoAux.Recordset.RecordCount > 0 Then
                AdoAux.Recordset.MoveFirst
                AdoAux.Recordset.Find ("Codigo = '" & Codigo & "' ")
                If Not AdoAux.Recordset.EOF Then
                   Opciones = Val(AdoAux.Recordset.Fields("SN"))
                   Real2 = AdoAux.Recordset.Fields("Valor_Hora")
                   Cantidad = AdoAux.Recordset.Fields("Horas_Sem") / 5
                   Real1 = Val(Hour(HoraDiaF - HoraDiaI) & "." & Minute(HoraDiaF - HoraDiaI))
                   If Opciones = 2 Then
                      If Real1 < Cantidad Then Real1 = Cantidad
                      If Real1 > Cantidad Then Real1 = Cantidad
                   End If
                   If HoraDiaF = HoraDiaI Then MsgBox "Iguales"
                   'MsgBox Real1
                   If Real1 <= 0 Then Real1 = 0.01
                   Real3 = Redondear(Real1 * Real2, 2)
                   SetAdoAddNew "Trans_Rol_Horas"
                   SetAdoFields "T", Normal
                   SetAdoFields "Codigo", Codigo
                   SetAdoFields "Fecha", Mifecha
                   SetAdoFields "Horas", Real1
                   SetAdoFields "Valor_Hora", Real2
                   SetAdoFields "Ing_Liquido", Real3
                   SetAdoFields "Item", NumEmpresa
                   SetAdoUpdate
                End If
             End If
   End If
  End With
End Sub

Private Sub Command1_Click()
 'Consultar
  RatonReloj
  Contador = 0
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  Select Case Modulo
    Case "FACTURAS": 'ActualizarVentasAnticipadas
    Case "ROL PAGOS": Actualizar_Entradas_Salidas
  End Select
  RatonNormal
  Unload AcEntSal
End Sub

Private Sub Command3_Click()
  Unload AcEntSal
End Sub

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Form_Load()
  Select Case Modulo
    Case "FACTURAS": AcEntSal.Caption = "ACTUALIZANDO VENTAS ANTICIPADAS"
    Case "ROL PAGOS": AcEntSal.Caption = "ACTUALIZANDO ENTRADAS/SALIDAS"
  End Select
  CentrarForm AcEntSal
  ConectarAdodc AdoAux
  ConectarAdodc AdoAux1
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

Function Calcular_Horas_Trabajadas(TiempoInicial As Single, _
                                   TiempoFinal As Single, _
                                   ValorHora As Single) As Single
Dim PC_Minutos As Integer
Dim PC_Horas As Integer
Dim Total_Horas As Single
Dim TiempoTrabajo As String
  TiempoTrabajo = CDate(TiempoFinal - TiempoInicial)
  PC_Minutos = Minute(TiempoTrabajo)
  PC_Horas = Hour(TiempoTrabajo)
  Total_Horas = PC_Minutos * (ValorHora / 60)
  Total_Horas = Redondear(Total_Horas + (Total_Horas * ValorHora), 2)
  Calcular_Horas_Trabajadas = Total_Horas
End Function

Public Sub Insertar_Sabado_Domingo()
'CheqSab
'CheqDom
End Sub
