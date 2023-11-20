VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form ResumenHoras 
   Caption         =   "RESUMEN DE HORAS LABORADAS"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Modulos"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir los resultados"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Resumen Horas"
            Object.ToolTipText     =   "Resumen horas laboradas"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Horas por Valor"
            Object.ToolTipText     =   "Resumen de horas laboradas por Valor"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Datos Rol"
            Object.ToolTipText     =   "Presenta Datos del Rol de Pagos"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCCostos 
      Bindings        =   "ResHorRo.frx":0000
      DataSource      =   "AdoCostos"
      Height          =   315
      Left            =   3675
      TabIndex        =   8
      Top             =   1365
      Visible         =   0   'False
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Costos"
   End
   Begin VB.CheckBox CheqCostos 
      Caption         =   "&Centro de Costos"
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
      Left            =   2100
      TabIndex        =   7
      Top             =   1365
      Width           =   1485
   End
   Begin VB.CheckBox CheqGrupo 
      Caption         =   "&Grupo de Empleados"
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
      Left            =   10395
      TabIndex        =   11
      Top             =   1050
      Width           =   2115
   End
   Begin VB.CheckBox CheqTipoProc 
      Caption         =   "&Tipo Proceso"
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
      Left            =   10395
      TabIndex        =   9
      Top             =   735
      Width           =   1485
   End
   Begin MSMask.MaskEdBox MBFechaF 
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   1155
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin VB.ComboBox CClientesF 
      Height          =   315
      Left            =   3675
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1050
      Visible         =   0   'False
      Width           =   6630
   End
   Begin VB.ComboBox CClientesI 
      Height          =   315
      Left            =   3675
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   735
      Visible         =   0   'False
      Width           =   6630
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "ResHorRo.frx":0018
      Height          =   4530
      Left            =   105
      TabIndex        =   14
      Top             =   1785
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox CheqPorEmp 
      Caption         =   "&Por Empleado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2100
      TabIndex        =   4
      Top             =   735
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&S"
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
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1890
      Width           =   330
   End
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   735
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "0"
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   315
      Top             =   2835
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
   Begin MSAdodcLib.Adodc AdoCostos 
      Height          =   330
      Left            =   315
      Top             =   3465
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
      Caption         =   "Costos"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   315
      Top             =   3150
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
      Caption         =   "Clientes"
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
      Left            =   105
      Top             =   6405
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   30
      CommandTimeout  =   30
      CursorType      =   2
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
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   315
      Top             =   3780
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
      Caption         =   "Grupo"
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
   Begin MSDataListLib.DataCombo DCGrupo 
      Bindings        =   "ResHorRo.frx":002F
      DataSource      =   "AdoGrupo"
      Height          =   315
      Left            =   12600
      TabIndex        =   12
      Top             =   1050
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Grupo"
   End
   Begin MSDataListLib.DataCombo DCTipoProc 
      Bindings        =   "ResHorRo.frx":0046
      DataSource      =   "AdoTipoProc"
      Height          =   315
      Left            =   12600
      TabIndex        =   10
      Top             =   735
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "TipoProc"
   End
   Begin MSAdodcLib.Adodc AdoTipoProc 
      Height          =   330
      Left            =   315
      Top             =   4095
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
      Caption         =   "TipoProc"
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Horas"
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
      Left            =   525
      TabIndex        =   16
      Top             =   7980
      Width           =   1170
   End
   Begin VB.Label LblHoras 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1680
      TabIndex        =   15
      Top             =   7980
      Width           =   1695
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
      Top             =   1155
      Width           =   750
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   10395
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResHorRo.frx":0060
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResHorRo.frx":037A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResHorRo.frx":0694
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResHorRo.frx":09AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ResHorRo.frx":0CC8
            Key             =   ""
         EndProperty
      EndProperty
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
      Top             =   735
      Width           =   750
   End
End
Attribute VB_Name = "ResumenHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TDiaLab(32) As String
Dim TValorDia(32) As Double
Dim VIng(16) As String
Dim VEgr(16) As String

Dim CliEmpI As String
Dim CliEmpF As String
Dim CodCostos As String
Dim CodDetalle As String
Dim CodTipoProc As String
Dim CodGrupo As String
'

Public Sub Resumen_Horas_por_Mes()
  sSQL = "SELECT C.Cliente,M.Dia_Mes As Mes,"
  For I = 1 To 31
      sSQL = sSQL & "BDM.D_" & Format(I, "00") & ","
  Next I
  sSQL = sSQL & "BDM.TOTAL,CR.Valor_Hora,(BDM.Total * CR.Valor_Hora) As Ing_Liquido " _
       & "FROM Clientes As C,Balances_Mes As BDM,Tabla_Dias_Meses As M,Catalogo_Rol_Pagos As CR " _
       & "WHERE BDM.TB = 'RPH' " _
       & "AND BDM.CodigoU = '" & CodigoUsuario & "' " _
       & "AND BDM.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' " _
       & "AND M.Tipo = 'M' " _
       & "AND C.Codigo = BDM.Codigo " _
       & "AND BDM.Item = CR.Item " _
       & "AND BDM.Codigo = CR.Codigo " _
       & "AND BDM.NoMes = M.No_D_M " _
       & "ORDER BY C.Cliente,BDM.NoMes "
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
  Total = 0
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          Total = Total + .Fields("TOTAL")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  
  LblHoras.Caption = Format(Total, "#,##0.00")
  ResumenHoras.Caption = "RESUMEN DE HORAS LABORADAS"
  DGQuery.Visible = True
  RatonNormal
  Opcion = 2
End Sub

Private Sub CheqCostos_Click()
    If CheqCostos.value <> 0 Then DCCostos.Visible = True Else DCCostos.Visible = False
End Sub

Private Sub CheqGrupo_Click()
    If CheqGrupo.value <> 0 Then DCGrupo.Visible = True Else DCGrupo.Visible = False
End Sub

Private Sub CheqPorEmp_Click()
    If CheqPorEmp.value <> 0 Then
       CClientesI.Visible = True
       CClientesF.Visible = True
    Else
       CClientesI.Visible = False
       CClientesF.Visible = False
    End If
End Sub

Private Sub CheqTipoProc_Click()
    If CheqTipoProc.value <> 0 Then DCTipoProc.Visible = True Else DCTipoProc.Visible = False
End Sub

Private Sub Command3_Click()
  Unload ResumenHoras
End Sub
'Imprimir Datos
Public Sub Impresiones()
  Select Case Opcion
    Case 0: Imprimir_Entradas_Salidas AdoQuery, 2, 6, 1, True
    Case 1: Imprimir_Entradas_Salidas AdoQuery, 2, 6, 2, True
    Case 2: Imprimir_Entradas_Salidas AdoQuery, 1, 9, 3
  End Select
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
   Keys_Especiales Shift
   If CtrlDown And KeyCode = vbKeyF1 Then GenerarDataTexto ResumenHoras, AdoQuery
End Sub

Private Sub Form_Activate()
  Opcion = 0
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  
  CClientesI.Clear
  CClientesF.Clear
  sSQL = "SELECT Cod_Rol_Pago, Detalle " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Tipo_Rubro = 'PER' " _
       & "GROUP BY Cod_Rol_Pago, Detalle " _
       & "ORDER BY Detalle, Cod_Rol_Pago "
  SelectDB_Combo DCTipoProc, AdoTipoProc, sSQL, "Detalle"
  
  sSQL = "SELECT Grupo_Rol " _
       & "FROM Catalogo_Rol_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "GROUP BY Grupo_Rol " _
       & "ORDER BY Grupo_Rol "
  SelectDB_Combo DCGrupo, AdoGrupo, sSQL, "Grupo_Rol"
  
  sSQL = "SELECT TS.Codigo, CS.Detalle " _
       & "FROM Catalogo_SubCtas As CS, Trans_SubCtas As TS " _
       & "WHERE CS.Item = '" & NumEmpresa & "' " _
       & "AND CS.Periodo = '" & Periodo_Contable & "' " _
       & "AND CS.Item = TS.Item " _
       & "AND CS.Periodo = TS.Periodo " _
       & "AND CS.Codigo = TS.Codigo " _
       & "GROUP BY TS.Codigo, CS.Detalle " _
       & "ORDER BY CS.Detalle "
  SelectDB_Combo DCCostos, AdoCostos, sSQL, "Detalle"
  
  sSQL = "SELECT C.Cliente, C.Codigo, CR.Salario " _
       & "FROM Clientes As C, Catalogo_Rol_Pagos As CR " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' " _
       & "AND CR.Codigo = C.Codigo " _
       & "ORDER BY Cliente "
  Select_Adodc AdoClientes, sSQL
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
       CClientesI.AddItem "Todos"
       CClientesF.AddItem "Todos"
       Do While Not .EOF
          CClientesI.AddItem .Fields("Cliente")
          CClientesF.AddItem .Fields("Cliente")
          CliEmpF = .Fields("Cliente")
         .MoveNext
       Loop
   Else
       CClientesI.AddItem "No existe Datos"
       CClientesF.AddItem "No existe Datos"
   End If
  End With
  CClientesI.Text = CClientesI.List(0)
  CClientesF.Text = CClientesI.List(0)
  
     AdoQuery.Visible = False
     DGQuery.Visible = False
     DGQuery.width = MDI_X_Max - 100
     DGQuery.Height = MDI_Y_Max - 2000
     
     AdoQuery.width = MDI_X_Max - 3000
     AdoQuery.Top = DGQuery.Top + DGQuery.Height + 10
     Label1.Top = AdoQuery.Top
     LblHoras.Top = AdoQuery.Top
     Label1.Left = AdoQuery.Left + AdoQuery.width + 10
     LblHoras.Left = Label1.Left + Label1.width
     
     AdoQuery.Visible = True
     DGQuery.Visible = True
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoQuery
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoCostos
  ConectarAdodc AdoTipoProc
  ConectarAdodc AdoClientes
End Sub

Private Sub MBFechaF_GotFocus()
  MarcarTexto MBFechaF
End Sub

Private Sub MBFechaF_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaF_LostFocus()
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  sSQL = "SELECT Cod_Rol_Pago, Detalle " _
       & "FROM Trans_Rol_de_Pagos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fecha_D >= #" & FechaIni & "# " _
       & "AND Fecha_H <= #" & FechaFin & "# " _
       & "AND Tipo_Rubro = 'PER' " _
       & "GROUP BY Cod_Rol_Pago, Detalle " _
       & "ORDER BY Detalle, Cod_Rol_Pago "
  SelectDB_Combo DCTipoProc, AdoTipoProc, sSQL, "Detalle"
End Sub

Private Sub MBFechaI_GotFocus()
  MarcarTexto MBFechaI
End Sub

Private Sub MBFechaI_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFechaI_LostFocus()
  FechaValida MBFechaI
  MBFechaF = UltimoDiaMes(MBFechaI)
End Sub

Public Sub Limpiar_Reporte_Rol()
  RatonReloj
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI.Text)
  FechaFin = BuscarFecha(MBFechaF.Text)
  SQL2 = "SELECT * " _
       & "FROM Balances_Mes " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TB = 'RPH' "
  Select_Adodc AdoAux, SQL2
  
  SQL2 = "DELETE * " _
       & "FROM Balances_Mes " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' " _
       & "AND TB = 'RPH' "
  Ejecutar_SQL_SP SQL2
  
  sSQL = "SELECT CR.*,C.Cliente " _
       & "FROM Clientes As C,Catalogo_Rol_Pagos As CR " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' "
  If CheqPorEmp.value <> 0 Then sSQL = sSQL & "AND C.Cliente BETWEEN '" & CliEmpI & "' AND '" & CliEmpF & "' "
  sSQL = sSQL & "AND C.Codigo = CR.Codigo " _
       & "ORDER BY C.Cliente,CR.Codigo "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount Then
       I = Month(MBFechaI)
       J = Month(MBFechaF)
       Contador = 0
       Do While Not .EOF
          NoMeses = Month(MBFechaI)
          Contador = Contador + 1
          ResumenHoras.Caption = Format(Contador / .RecordCount, "00%") & " - " & .Fields("Cliente")
          For K = 1 To J - I + 1
              SetAdoAddNew "Balances_Mes"
              SetAdoFields "Codigo", .Fields("Codigo")
              SetAdoFields "TB", "RPH"
              SetAdoFields "NoMes", CByte(NoMeses)
              SetAdoFields "Item", NumEmpresa
              SetAdoFields "CodigoU", CodigoUsuario
              SetAdoUpdate
              NoMeses = NoMeses + 1
              If NoMeses > 12 Then NoMeses = 1
          Next K
         .MoveNext
       Loop
   End If
  End With
  sSQL = "SELECT * " _
       & "FROM Balances_Mes " _
       & "WHERE CodigoU = '" & CodigoUsuario & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & " " _
       & "AND TB = 'RPH' " _
       & "ORDER BY Codigo,NoMes "
  Select_Adodc AdoAux, sSQL
  RatonNormal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
 'MsgBox Button.key
  FechaValida MBFechaI
  FechaValida MBFechaF
  FechaIni = BuscarFecha(MBFechaI)
  FechaFin = BuscarFecha(MBFechaF)
  
  CliEmpI = Ninguno
  CliEmpF = Ninguno
  With AdoClientes.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & CClientesI.Text & "' ")
       If Not .EOF Then CliEmpI = .Fields("Cliente")
      .MoveFirst
      .Find ("Cliente = '" & CClientesF.Text & "' ")
       If Not .EOF Then CliEmpF = .Fields("Cliente")
   End If
  End With
  
  CodCostos = Ninguno
  With AdoCostos.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Detalle = '" & DCCostos.Text & "' ")
       If Not .EOF Then CodCostos = .Fields("Codigo")
   End If
  End With
  
  CodDetalle = Ninguno
  With AdoTipoProc.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Detalle = '" & DCTipoProc.Text & "' ")
       If Not .EOF Then CodDetalle = .Fields("Cod_Rol_Pago")
   End If
  End With
  
  CodGrupo = Ninguno
  With AdoGrupo.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Grupo_Rol = '" & DCGrupo.Text & "' ")
       If Not .EOF Then CodGrupo = .Fields("Grupo_Rol")
   End If
  End With
  
  Select Case Button.key
    Case "Salir"
         Unload ResumenHoras
    Case "Imprimir"
         Impresiones
    Case "Resumen Horas"
         Resumen_Horas
    Case "Resumen Horas por Mes"
         Resumen_Horas_por_Mes
    Case "Datos Rol"
         Datos_Rol
    Case ""
  End Select
End Sub

Public Sub Resumen_Horas()
  DGQuery.Visible = False
  Limpiar_Reporte_Rol
  RatonReloj
 'Transmitimos Entradas/Salidas de horas trabajadas
  sSQL = "SELECT Codigo,Fecha,Horas,Valor_Hora " _
       & "FROM Trans_Rol_Horas " _
       & "WHERE Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If CheqPorEmp.value <> 0 Then sSQL = sSQL & "AND C.Cliente BETWEEN '" & CliEmpI & "' AND '" & CliEmpF & "' "
  sSQL = sSQL & "ORDER BY Codigo,Fecha "
  Select_Adodc AdoAux, sSQL
  With AdoAux.Recordset
   If .RecordCount > 0 Then
       Codigo = .Fields("Codigo")
       NoMeses = Month(.Fields("Fecha"))
       Mifecha = .Fields("Fecha")
       NoAnio = Year(.Fields("Fecha"))
       For I = 1 To 31
           TValorDia(I) = 0
           TDiaLab(I) = Ninguno
       Next I
       Contador = 0
       Do While Not .EOF
          Contador = Contador + 1
          ResumenHoras.Caption = Format(Contador / .RecordCount, "00%")
          If Codigo <> .Fields("Codigo") Or NoMeses <> Month(.Fields("Fecha")) Then
             Total = 0: J = 0
             SQL2 = "UPDATE Balances_Mes " _
                  & "SET "
             For I = 1 To 31
                 SQL2 = SQL2 & "D" & Format(I, "00") & " = '" & TDiaLab(I) & "'," _
                      & "D_" & Format(I, "00") & " = " & TValorDia(I) & ", "
                 If TDiaLab(I) <> Ninguno Then J = J + 1
                 Total = Total + TValorDia(I)
             Next I
             SQL2 = SQL2 & "Dias = " & J & ", Total = " & Total & " " _
                  & "WHERE Item = '" & NumEmpresa & "' " _
                  & "AND CodigoU = '" & CodigoUsuario & "' " _
                  & "AND Codigo = '" & Codigo & "' " _
                  & "AND NoMes = " & NoMeses & " " _
                  & "AND TB = 'RPH' "
             'MsgBox SQL2
             Ejecutar_SQL_SP SQL2
             For I = 1 To 31
                 TValorDia(I) = 0
                 TDiaLab(I) = Ninguno
             Next I
             Codigo = .Fields("Codigo")
             NoMeses = Month(.Fields("Fecha"))
             Mifecha = .Fields("Fecha")
             NoAnio = Year(.Fields("Fecha"))
          End If
          I = Day(.Fields("Fecha"))
          TValorDia(I) = TValorDia(I) + .Fields("Horas")
          If TValorDia(I) > 0 Then TDiaLab(I) = "X"
         .MoveNext
       Loop
   End If
  End With
  DGQuery.Caption = "RESUMEN DE HORAS LABORADAS DEL " & MBFechaI.Text & " AL " & MBFechaF.Text
  'If SQL_Server = False Then MsgBox "Proceso Terminado, Puede Consultar"
  sSQL = "SELECT C.Cliente,M.Dia_Mes As Mes,"
  For I = 1 To 31
      sSQL = sSQL & "BDM.D" & Format(I, "00") & ","
  Next I
  sSQL = sSQL & "BDM.Dias,CR.Valor_Hora,(BDM.Total * CR.Valor_Hora) As Ing_Liquido " _
       & "FROM Clientes As C,Balances_Mes As BDM,Tabla_Dias_Meses As M,Catalogo_Rol_Pagos As CR " _
       & "WHERE BDM.TB = 'RPH' " _
       & "AND BDM.CodigoU = '" & CodigoUsuario & "' " _
       & "AND BDM.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' " _
       & "AND M.Tipo = 'M' " _
       & "AND C.Codigo = BDM.Codigo " _
       & "AND BDM.Codigo = CR.Codigo " _
       & "AND BDM.Item = CR.Item " _
       & "AND BDM.NoMes = M.No_D_M " _
       & "ORDER BY C.Cliente,BDM.NoMes "
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
  Opcion = 1
End Sub

Public Sub X()
  sSQL = "SELECT TES.Fecha,TES.Hora,C.Cliente,TES.Tarea " _
       & "FROM Trans_Entrada_Salida As TES, Clientes As C " _
       & "WHERE TES.Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
       & "AND TES.ES = 'H' " _
       & "AND TES.Item = '" & NumEmpresa & "' " _
       & "AND TES.Periodo = '" & Periodo_Contable & "' "
  If CheqPorEmp.value <> 0 Then sSQL = sSQL & "AND C.Cliente BETWEEN '" & CliEmpI & "' AND '" & CliEmpF & "' "
  sSQL = sSQL & "AND TES.Codigo = C.Codigo " _
       & "ORDER BY C.Cliente,TES.Fecha,TES.Hora "
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
  DGQuery.Caption = "RESUMEN DE HORAS POR TAREAS"
  Opcion = 4
End Sub

Public Function Insertar_Ing_Egr(TipoProc As String, NombreCampo As String) As Long
Dim IdResul As Long
Dim NoEsta As Boolean
 IdResul = 0
 If NombreCampo <> "Neto_Recibir" Then
    If TipoProc = "Ing" Then
       NoEsta = True
       For K = 1 To 15
           If NombreCampo = VIng(K) Then
              NoEsta = False
              Exit For
           End If
       Next K
       If NoEsta Then
          For K = 1 To 15
              If VIng(K) = "" Then
                 VIng(K) = NombreCampo
                 Exit For
              End If
          Next K
       End If
       For K = 1 To 15
           If NombreCampo = VIng(K) Then
              IdResul = K
              Exit For
           End If
       Next K
    ElseIf TipoProc = "Egr" Then
       NoEsta = True
       For K = 1 To 15
           If NombreCampo = VEgr(K) Then
              NoEsta = False
              Exit For
           End If
       Next K
       If NoEsta Then
          For K = 1 To 15
              If VEgr(K) = "" Then
                 VEgr(K) = NombreCampo
                 IdResul = K
                 Exit For
              End If
          Next K
       End If
       For K = 1 To 15
           If NombreCampo = VEgr(K) Then
              IdResul = K
              Exit For
           End If
       Next K
    End If
 End If
 Insertar_Ing_Egr = IdResul
End Function

Public Sub Datos_Rol()
Dim Neto_Rec As Boolean
    Neto_Rec = False
    For K = 1 To 15
        VIng(K) = ""
        VEgr(K) = ""
    Next K

   'Borrarmos la tabla temporal del rol
    SQL1 = "DELETE * " _
         & "FROM Asiento_Rol_Colectivo " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP SQL1
    SQL1 = ""
    
    sSQL = "SELECT CR.Fecha, RP.Fecha_H, C.Cliente, RP.Detalle, RP.Ingresos, RP.Egresos, CR.Grupo_Rol, CR.Cta_Transferencia, " _
         & "C.CI_RUC, C.Codigo, CR.IEESS_Per, CR.IEESS_Pat, CR.Pagar_Fondo_Reserva, CR.Salario, RP.Cod_Rol_Pago, RP.ID " _
         & "FROM Trans_Rol_de_Pagos As RP, Clientes As C, Catalogo_Rol_Pagos As CR " _
         & "WHERE RP.Fecha_D >= #" & FechaIni & "# " _
         & "AND RP.Fecha_H <= #" & FechaFin & "# " _
         & "AND RP.Tipo_Rubro = 'PER' " _
         & "AND RP.Item = '" & NumEmpresa & "' " _
         & "AND RP.Periodo = '" & Periodo_Contable & "' "
    If CheqPorEmp.value <> 0 Then sSQL = sSQL & "AND C.Cliente BETWEEN '" & CliEmpI & "' AND '" & CliEmpF & "' "
    If CheqCostos.value <> 0 Then sSQL = sSQL & "AND CR.SubModulo = '" & CodCostos & "' "
    If CheqTipoProc.value <> 0 Then sSQL = sSQL & "AND RP.Cod_Rol_Pago = '" & CodDetalle & "' "
    If CheqGrupo.value <> 0 Then sSQL = sSQL & "AND CR.Grupo_Rol = '" & CodGrupo & "' "
    sSQL = sSQL _
         & "AND RP.Codigo = C.Codigo " _
         & "AND CR.Codigo = C.Codigo " _
         & "AND RP.Item = CR.Item " _
         & "AND RP.Periodo = CR.Periodo " _
         & "ORDER BY C.Cliente, RP.Fecha_H, CR.Fecha, RP.ID "
    Select_Adodc AdoAux, sSQL
   'Insertamos los Empleados del rol
    Contador = 1
    With AdoAux.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         I = 1
         J = 1
         CodigoCli = .Fields("Codigo")
         CodigoB = .Fields("CI_RUC")
         CodigoA = .Fields("Grupo_Rol")
         Mifecha = .Fields("Fecha")
         FechaTexto = .Fields("Fecha_H")
         NombreCliente = .Fields("Cliente")
         Si_No = .Fields("Pagar_Fondo_Reserva")
         Real1 = .Fields("IEESS_Pat")
         Real2 = .Fields("IEESS_Per")
         
         SetAdoAddNew "Asiento_Rol_Colectivo"
         SetAdoFields "No_", Format(Contador, "000")
         SetAdoFields "Codigo", CodigoCli
         SetAdoFields "C_I", CodigoB
         SetAdoFields "Nombre_Empleado", NombreCliente
         SetAdoFields "Grupo_Rol", CodigoA
         SetAdoFields "Fecha", .Fields("Fecha_H")
         SetAdoFields "Porc_Apo_Pat", Real1
         SetAdoFields "Porc_Apo_Per", Real2
         SetAdoFields "Item", NumEmpresa
         SetAdoFields "CodigoU", CodigoUsuario
         SetAdoFields "Fecha_Ing", Mifecha
         SetAdoFields "FR", Si_No
         SetAdoFields "I", "|"
         SetAdoFields "II", "|"
         
        'Porc_Apo_Pat, Porc_Apo_Per
         SQL1 = "No_, C_I, Nombre_Empleado, Grupo_Rol, Fecha, Fecha_Ing, FR, I, "
         Do While Not .EOF
            If CodigoCli <> .Fields("Codigo") Or (FechaTexto <> .Fields("Fecha_H")) Then
               SetAdoUpdate
               I = 1
               J = 1
               Contador = Contador + 1
               CodigoCli = .Fields("Codigo")
               CodigoB = .Fields("CI_RUC")
               CodigoA = .Fields("Grupo_Rol")
               Mifecha = .Fields("Fecha")
               FechaTexto = .Fields("Fecha_H")
               NombreCliente = .Fields("Cliente")
               Si_No = .Fields("Pagar_Fondo_Reserva")
               Real1 = .Fields("IEESS_Pat")
               Real2 = .Fields("IEESS_Per")
               
               SetAdoAddNew "Asiento_Rol_Colectivo"
               SetAdoFields "No_", Format(Contador, "000")
               SetAdoFields "Codigo", CodigoCli
               SetAdoFields "C_I", CodigoB
               SetAdoFields "Nombre_Empleado", NombreCliente
               SetAdoFields "Grupo_Rol", CodigoA
               SetAdoFields "Fecha", .Fields("Fecha_H")
               SetAdoFields "Porc_Apo_Pat", Real1
               SetAdoFields "Porc_Apo_Per", Real2
               SetAdoFields "Item", NumEmpresa
               SetAdoFields "CodigoU", CodigoUsuario
               SetAdoFields "Fecha_Ing", Mifecha
               SetAdoFields "FR", Si_No
               SetAdoFields "I", "|"
               SetAdoFields "II", "|"
            End If
            If .Fields("Ingresos") > 0 Then
                I = Insertar_Ing_Egr("Ing", .Fields("Cod_Rol_Pago"))
                If I > 0 Then SetAdoFields "Ing_" & Format(I, "00"), .Fields("Ingresos")
            End If
            If .Fields("Egresos") > 0 Then
                
                J = Insertar_Ing_Egr("Egr", .Fields("Cod_Rol_Pago"))
                If J > 0 Then SetAdoFields "Egr_" & Format(J, "00"), .Fields("Egresos")
            End If
            If .Fields("Cod_Rol_Pago") = "Neto_Recibir" Then
                SetAdoFields "Neto_Recibir", .Fields("Egresos")
                Neto_Rec = True
            End If
           .MoveNext
         Loop
         SetAdoUpdate
     End If
    End With
''    Cadena = ""
''    For I = 1 To 15
''        Cadena = Cadena & "Egr_" & Format(I, "00") & " As " & VEgr(I) & vbCrLf
''    Next I
''    MsgBox Cadena
    For I = 1 To 15
        If VIng(I) <> "" Then SQL1 = SQL1 & "Ing_" & Format(I, "00") & " As " & VIng(I) & ", "
    Next I
    SQL1 = SQL1 & "II, "
    For I = 1 To 15
        If VEgr(I) <> "" Then SQL1 = SQL1 & "Egr_" & Format(I, "00") & " As " & VEgr(I) & ", "
    Next I
    If Neto_Rec Then
       SQL1 = SQL1 & " Neto_Recibir"
    Else
       SQL1 = TrimStrg(SQL1)
       K = Len(SQL1)
       SQL1 = MidStrg(SQL1, 1, K - 1)
    End If
    
    sSQL = "SELECT " & SQL1 & " " _
         & "FROM Asiento_Rol_Colectivo " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY No_, Nombre_Empleado, Fecha "
    Select_Adodc_Grid DGQuery, AdoQuery, sSQL, , True
    DGQuery.Caption = "RESUMEN DE NOMINA"
    Opcion = 3
End Sub

