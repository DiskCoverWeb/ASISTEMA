VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FListarEmpleados 
   Caption         =   "LISTA DE EMPLEADOS"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar TBarCliente 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Todos"
            Object.ToolTipText     =   "Todos los Empleados"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SinCta"
            Object.ToolTipText     =   "Listar sin Acreditar a la Cuenta Bancaria"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ConCta"
            Object.ToolTipText     =   "Listar con Acreditacion a la Cuenta Bancaria"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Retirados"
            Object.ToolTipText     =   "Lista de Empleados Retirados o Anulados"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SRI"
            Object.ToolTipText     =   "SRI"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Varones"
            Object.ToolTipText     =   "Listar Solo Varones"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Mujeres"
            Object.ToolTipText     =   "Listar Solo Mujeres"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ingresos_Emp"
            Object.ToolTipText     =   "Ingresos Adicionales Empleados"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Egresos_Emp"
            Object.ToolTipText     =   "Egresos Adicionales Empleados"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Grabar_I_E"
            Object.ToolTipText     =   "Grabar Ingresos/Egresos al Rol de Pagos"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   7665
         TabIndex        =   4
         Top             =   0
         Width           =   4110
         Begin MSDataListLib.DataCombo DCMes 
            Bindings        =   "FLstEmpl.frx":0000
            DataSource      =   "AdoMes"
            Height          =   315
            Left            =   2730
            TabIndex        =   7
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "Diciembre"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.ComboBox CAnio 
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
            Left            =   1890
            TabIndex        =   5
            Text            =   "2000"
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FECHA: (Ano/Mes)"
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
            TabIndex        =   6
            Top             =   210
            Width           =   1800
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "INGRESOS/EGRESOS/OTROS EMPLEADOS"
      TabPicture(0)   =   "FLstEmpl.frx":0015
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "AdoIEEmpleados"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DGIEEmpleados"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "LISTA DE NOMINA DE EMPLEADOS"
      TabPicture(1)   =   "FLstEmpl.frx":0031
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGListRol"
      Tab(1).Control(1)=   "AdoListRol"
      Tab(1).ControlCount=   2
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
         Left            =   3255
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4620
         Width           =   330
      End
      Begin MSDataGridLib.DataGrid DGListRol 
         Bindings        =   "FLstEmpl.frx":004D
         Height          =   4005
         Left            =   -74895
         TabIndex        =   1
         Top             =   420
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   7064
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LISTA DE EMPLEADOS"
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSAdodcLib.Adodc AdoListRol 
         Height          =   330
         Left            =   -74895
         Top             =   4620
         Width           =   3165
         _ExtentX        =   5583
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
      Begin MSDataGridLib.DataGrid DGIEEmpleados 
         Bindings        =   "FLstEmpl.frx":0066
         Height          =   4215
         Left            =   105
         TabIndex        =   2
         ToolTipText     =   "<Ctrl+Insert> Llama un dato del mes anterior, <Ctrl+Delete> Limpia los datos del Campo"
         Top             =   420
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LISTA DE EMPLEADOS"
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSAdodcLib.Adodc AdoIEEmpleados 
         Height          =   330
         Left            =   105
         Top             =   4620
         Width           =   3165
         _ExtentX        =   5583
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
         Caption         =   "IEEmpleados"
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
   End
   Begin MSAdodcLib.Adodc AdoSRI 
      Height          =   330
      Left            =   210
      Top             =   1575
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
      Caption         =   "SRI"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   1995
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
   Begin MSAdodcLib.Adodc AdoMes 
      Height          =   330
      Left            =   210
      Top             =   2310
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
      Caption         =   "Mes"
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   12075
      Top             =   735
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
            Picture         =   "FLstEmpl.frx":0083
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":039D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":06B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":09D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":0CEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":1005
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":5F8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":62A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":65C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":104C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":1A2FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FLstEmpl.frx":1A619
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FListarEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FechaRep As String
   
Private Sub CAnio_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub Command3_Click()
  Unload FListarEmpleados
End Sub

Private Sub DCMes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCMes_LostFocus()
  AdoMes.Recordset.MoveFirst
  AdoMes.Recordset.Find ("Dia_Mes = '" & DCMes.Text & "' ")
  If Not AdoMes.Recordset.EOF Then
     NoMes = AdoMes.Recordset.fields("No_D_M")
  Else
     NoMes = Month(FechaSistema)
  End If
  FechaRep = "01/" & Format$(NoMes, "00") & "/" & CAnio.Text
  FechaRep = UltimoDiaMes(FechaRep)
  
  Datos_IESS FechaRep
End Sub

Private Sub DGIEEmpleados_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FechaI As String
Dim FechaF As String
Dim MesRP As Integer
Dim AnioRP As Integer
    
    Codigo1 = Ninguno
    Codigo2 = Ninguno
    If DGIEEmpleados.Col >= 0 Then
        If (DGIEEmpleados.Columns.Count - 2) > 0 Then
            Codigo1 = DGIEEmpleados.Columns.Item(DGIEEmpleados.Col).Caption        ' I/E
            Codigo2 = DGIEEmpleados.Columns.Item(DGIEEmpleados.Columns.Count - 2)  ' Nombre del Campo
        End If
    End If
    Codigo3 = "I"
    Select Case Codigo2
      Case "I": Codigo3 = "Ingresos"
      Case "E": Codigo3 = "Egresos"
    End Select
    
    AnioRP = Year(FechaRep)
    MesRP = Month(FechaRep) - 1
    If MesRP < 0 Then
       MesRP = 12
       AnioRP = AnioRP - 1
    End If
    FechaI = "01/" & Format(MesRP, "00") & "/" & Format(AnioRP, "0000")
    FechaF = UltimoDiaMes(FechaI)
    FechaI = BuscarFecha(FechaI)
    FechaF = BuscarFecha(FechaF)

  Keys_Especiales Shift
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGIEEmpleados.Visible = False
     GenerarDataTexto FListarEmpleados, AdoIEEmpleados, True
     DGIEEmpleados.Visible = True
  End If
  If KeyCode = vbKeyReturn Then
     If AdoIEEmpleados.Recordset.RecordCount > 0 Then
        AdoIEEmpleados.Recordset.MoveNext
        If AdoIEEmpleados.Recordset.EOF Then AdoIEEmpleados.Recordset.MoveFirst
     End If
  End If
  If CtrlDown And KeyCode = vbKeyDelete Then
     If AdoIEEmpleados.Recordset.RecordCount > 0 Then
        Codigo3 = "I"
        If DGIEEmpleados.Col >= 1 Then
           sSQL = "UPDATE Asiento_RP_IE " _
                & "SET " & Codigo1 & " = 0 " _
                & "WHERE Item = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
        End If
       'Ingresos_Egresos_Empleados Codigo2
        sSQL = "SELECT * " _
             & "FROM Asiento_RP_IE " _
             & "ORDER BY Empleado "
        Select_Adodc_Grid DGIEEmpleados, AdoIEEmpleados, sSQL
        MsgBox "Proceso terminado"
        DCMes.SetFocus
     End If
  End If
  If CtrlDown And KeyCode = vbKeyInsert Then
     If AdoIEEmpleados.Recordset.RecordCount > 0 Then
        
        If DGIEEmpleados.Col >= 1 Then
           If SQL_Server Then
              sSQL = "UPDATE Asiento_RP_IE " _
                   & "SET " & Codigo1 & " = TB." & Codigo3 & " " _
                   & "FROM Asiento_RP_IE As T, Trans_Rol_de_Pagos As TB "
           Else
              sSQL = "UPDATE Asiento_RP_IE As T, Trans_Rol_de_Pagos As TB " _
                   & "SET T." & Codigo1 & " = TB." & Codigo3 & " "
           End If
           sSQL = sSQL _
                & "WHERE TB.Item = '" & NumEmpresa & "' " _
                & "AND TB.Periodo = '" & Periodo_Contable & "' " _
                & "AND TB.Fecha_D >= #" & FechaI & "# " _
                & "AND TB.Fecha_H <=#" & FechaF & "# " _
                & "AND TB.Cod_Rol_Pago = '" & Codigo1 & "' " _
                & "AND TB.Codigo = T.Codigo " _
                & "AND TB.Item = T.Item "
           Ejecutar_SQL_SP sSQL
        End If
       'Ingresos_Egresos_Empleados Codigo2
        sSQL = "SELECT * " _
             & "FROM Asiento_RP_IE " _
             & "ORDER BY Empleado "
        Select_Adodc_Grid DGIEEmpleados, AdoIEEmpleados, sSQL
        MsgBox "Proceso terminado"
        DCMes.SetFocus
     End If
  End If
End Sub

Private Sub DGListRol_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If AltDown And KeyCode = vbKeyS Then Unload Me
  
  If CtrlDown And KeyCode = vbKeyF1 Then
     DGListRol.Visible = False
     GenerarDataTexto FListarEmpleados, AdoListRol, True
     DGListRol.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyP Then
     DGListRol.Visible = False
     ImprimirAdo AdoListRol, True, 1, 9
     DGListRol.Visible = True
  End If
End Sub

Private Sub Form_Activate()
  SSTab1.Tab = 1
  SSTab1.Height = MDI_Y_Max - SSTab1.Top - 50
  SSTab1.width = MDI_X_Max - SSTab1.Left - 50

  DGListRol.Height = SSTab1.Height - DGListRol.Top - 450
  DGListRol.width = SSTab1.width - 200
  AdoListRol.Top = DGListRol.Top + DGListRol.Height + 10
  
  DGIEEmpleados.Height = SSTab1.Height - DGIEEmpleados.Top - 450
  DGIEEmpleados.width = SSTab1.width - 200
  AdoIEEmpleados.Top = DGIEEmpleados.Top + DGIEEmpleados.Height + 10
  Command3.Top = DGIEEmpleados.Top + DGIEEmpleados.Height + 10
  Command3.Left = AdoIEEmpleados.Left + AdoIEEmpleados.width
  
  SSTab1.Tab = 0
  
  sSQL = "SELECT Dia_Mes,No_D_M " _
      & "FROM Tabla_Dias_Meses " _
      & "WHERE Tipo = 'M' " _
      & "AND No_D_M BETWEEN 1 and 12 " _
      & "ORDER BY No_D_M "
  SelectDB_Combo DCMes, AdoMes, sSQL, "Dia_Mes"
  DCMes.Text = MesesLetras(Month(FechaSistema))
  
  CAnio.Clear
  For I = 2000 To Year(FechaSistema)
      CAnio.AddItem CStr(I)
  Next I
  CAnio.Text = Year(FechaSistema)
  ListaEmpleadosRolPagos
  RatonNormal
  CAnio.SetFocus
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoListRol
  ConectarAdodc AdoIEEmpleados
  ConectarAdodc AdoSRI
  ConectarAdodc AdoAux
  ConectarAdodc AdoMes
End Sub

Private Sub ListaEmpleadosRolPagos(Optional TipoList As Byte)
  sSQL = "SELECT CR.T,C.Sexo,C.Cliente As Empleado,C.Fecha_N,C.CI_RUC As Cedula,C.TD,CR.Fecha,CR.Grupo_Rol,CR.FechaVI,CR.FechaVF,C.Email," _
       & "C.Email2,CR.Salario,CR.SN,C.Profesion,C.Actividad,CR.Cta_Transferencia,CR.TC,CR.FP,CR.SubModulo,CR.Pagar_Fondo_Reserva," _
       & "CRC.Cta_Sueldo,CRC.Cta_Vacacion,CRC.Cta_Horas_Ext,CR.CodProfesion,CRC.Cta_IESS_Patronal,CRC.Cta_IESS_Personal,CRC.Cta_Aporte_Patronal_G," _
       & "CRC.Cta_Decimo_Cuarto_G,CRC.Cta_Decimo_Tercer_G,CRC.Cta_Fondo_Reserva_G,CRC.Cta_Vacaciones_G," _
       & "CRC.Cta_Decimo_Cuarto_P,CRC.Cta_Decimo_Tercer_P,CRC.Cta_Fondo_Reserva_P,CRC.Cta_Vacaciones_P," _
       & "CRC.Cta_Quincena,CR.Cta_Forma_Pago,CR.Pagar_Fondo_Reserva,CR.Codigo " _
       & "FROM Clientes As C, Catalogo_Rol_Pagos As CR, Catalogo_Rol_Cuentas As CRC " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND CR.Periodo = '" & Periodo_Contable & "' "
  If TipoList <= 4 Then
     Select Case TipoList
       Case 1: sSQL = sSQL & "AND CR.Cta_Transferencia = '" & Ninguno & "' "
       Case 2: sSQL = sSQL & "AND CR.Cta_Transferencia <> '" & Ninguno & "' "
       Case 3: sSQL = sSQL & "AND C.Sexo = 'M' "
       Case 4: sSQL = sSQL & "AND C.Sexo = 'F' "
     End Select
     sSQL = sSQL & "AND CR.T = '" & Normal & "' "
  Else
     sSQL = sSQL & "AND CR.T <> '" & Normal & "' "
  End If
  sSQL = sSQL _
       & "AND C.Codigo = CR.Codigo " _
       & "AND CR.Grupo_Rol = CRC.Grupo_Rol " _
       & "AND CR.Item = CRC.Item " _
       & "AND CR.Periodo = CRC.Periodo " _
       & "ORDER BY CR.T,C.Sexo,CR.Grupo_Rol,C.Cliente "
  Select_Adodc_Grid DGListRol, AdoListRol, sSQL
End Sub

Private Sub TBarCliente_ButtonClick(ByVal Button As ComctlLib.Button)
 'MsgBox Button.key
  Select Case Button.key
    Case "Todos": ListaEmpleadosRolPagos
    Case "SinCta": ListaEmpleadosRolPagos 1
    Case "ConCta": ListaEmpleadosRolPagos 2
    Case "Imprimir": ImprimirAdo AdoListRol, True, 1, 9
    Case "SRI": Ingresar_RDEP
    Case "Varones": ListaEmpleadosRolPagos 3
    Case "Mujeres": ListaEmpleadosRolPagos 4
    Case "Retirados": ListaEmpleadosRolPagos 5
    Case "Ingresos_Emp": Ingresos_Egresos_Empleados "I"
    Case "Egresos_Emp": Ingresos_Egresos_Empleados "E"
    Case "Grabar_I_E": Grabar_Ingresos_Egresos_Empleados
    Case "Salir": Unload Me
  End Select
End Sub

Public Sub Ingresar_RDEP()
Dim Anio As Integer
  Mensajes = "Desea Grabar los Empleados S/N? "
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then
      
     Anio = Year(FechaRep)
     FechaIni = BuscarFecha("01/01/" & Anio)
     FechaFin = BuscarFecha(FechaRep)
     
     NumRet = Maximo_De("Trans_Dependencia", "IDT")
     SQL2 = "DELETE * " _
          & "FROM Trans_Dependencia " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND AñoRet = " & Anio & " "
     Ejecutar_SQL_SP SQL2
     
     sSQL = "SELECT COUNT(Codigo) As CantMes,Codigo,SUM(InLiquido) As TInLiquido,SUM(InHorasExt) As TInHorasExt,SUM(Aporte_Per) As TAporte_Per,SUM(Neto_Recibir) As TNeto_Recibir " _
          & "FROM Trans_Rol_Pagos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Fecha BETWEEN #" & FechaIni & "# and #" & FechaFin & "# " _
          & "GROUP BY Codigo " _
          & "ORDER BY Codigo "
     Select_Adodc AdoListRol, sSQL
     'MsgBox sSQL
     RatonReloj
     With AdoListRol.Recordset
      If .RecordCount Then
          Do While Not .EOF
             SubTotal = .fields("TInLiquido") - .fields("TAporte_Per")
             SetAdoAddNew "Trans_Dependencia"
             SetAdoFields "IdRet", .fields("Codigo")
             SetAdoFields "SisSalNet", "2"
             SetAdoFields "SuelSal", .fields("TInLiquido")
             SetAdoFields "SobSuelComRemu", .fields("TInHorasExt")
             SetAdoFields "DecimTer", Redondear(.fields("TInLiquido") / 12, 2)
             SetAdoFields "DecimCuar", Redondear((Sueldo_Basico / 12) * .fields("CantMes"), 2)
             SetAdoFields "ApoPerIess", .fields("TAporte_Per")
             SetAdoFields "SubTotal", SubTotal
             SetAdoFields "BasImp", SubTotal
             SetAdoFields "AñoRet", Anio
             SetAdoFields "Fecha", FechaRep
             SetAdoFields "Serie", "001001"
             SetAdoFields "Autorizacion", "1234567890"
             SetAdoFields "IDT", NumRet
             SetAdoFields "CodigoU", CodigoUsuario
             SetAdoFields "Item", NumEmpresa
             SetAdoFields "Numero", .fields("CantMes")
             SetAdoUpdate
             NumRet = NumRet + 1
            .MoveNext
          Loop
       End If
     End With
     RatonNormal
  End If
End Sub

Public Sub Ingresos_Egresos_Empleados(I_E_Emp As String)
Dim AdoCon1 As ADODB.Connection
Dim RstSchema As ADODB.Recordset
Dim Tabla_ya_Existe As Boolean

   'Fecha del Rol Mesual
    AdoMes.Recordset.MoveFirst
    AdoMes.Recordset.Find ("Dia_Mes = '" & DCMes.Text & "' ")
    If Not AdoMes.Recordset.EOF Then
       NoMes = AdoMes.Recordset.fields("No_D_M")
    Else
       NoMes = Month(FechaSistema)
    End If
    FechaRep = "01/" & Format$(NoMes, "00") & "/" & CAnio.Text
    FechaRep = UltimoDiaMes(FechaRep)

   'Consultamos las cuentas de la tabla
    sSQL = "SELECT Codigo,Cuenta,I_E_Emp,Con_IESS,Cod_Rol_Pago " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND I_E_Emp = '" & I_E_Emp & "' " _
         & "AND DG = 'D' " _
         & "AND LEN(Cod_Rol_Pago) > 1 " _
         & "ORDER BY Codigo "
    Select_Adodc AdoSRI, sSQL
   
    RatonReloj
    Tabla_ya_Existe = False
    Set AdoCon1 = New ADODB.Connection
    AdoCon1.open AdoStrCnn
    Set RstSchema = AdoCon1.OpenSchema(adSchemaTables)
    Do Until RstSchema.EOF
       If RstSchema!TABLE_TYPE = "TABLE" And MidStrg(RstSchema!TABLE_NAME, 1, 1) <> "~" Then
          If RstSchema!TABLE_NAME = "Asiento_RP_IE" Then Tabla_ya_Existe = True
       End If
       RstSchema.MoveNext
    Loop
    AdoCon1.Close
    If Tabla_ya_Existe Then
       sSQL = "DROP TABLE Asiento_RP_IE "
       Ejecutar_SQL_SP sSQL
    End If

    sSQL = "CREATE TABLE Asiento_RP_IE (" _
         & "Codigo NVARCHAR(10) NULL," _
         & "Empleado NVARCHAR(60) NULL," _
         & "SUELDO FLOAT NULL," _
         & "HORAS_EXT FLOAT NULL,"
    With AdoSRI.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            sSQL = sSQL & CStr(.fields("Cod_Rol_Pago")) & " FLOAT NULL,"
           .MoveNext
         Loop
     End If
    End With
    sSQL = sSQL _
         & "I_E_Emp NVARCHAR(1) NULL," _
         & "Item NVARCHAR(3) NULL); "
    Ejecutar_SQL_SP sSQL

    sSQL = "INSERT INTO Asiento_RP_IE(Codigo,Empleado,SUELDO,I_E_Emp,Item) " _
         & "SELECT C.Codigo,C.Cliente,CRP.Salario,'" & I_E_Emp & "' As IE,'" & NumEmpresa & "' As Item " _
         & "FROM Clientes As C,Catalogo_Rol_Pagos AS CRP " _
         & "WHERE CRP.Item = '" & NumEmpresa & "' " _
         & "AND CRP.Periodo = '" & Periodo_Contable & "' " _
         & "AND CRP.T = '" & Normal & "' " _
         & "AND C.Codigo = CRP.Codigo "
    Ejecutar_SQL_SP sSQL
    With AdoSRI.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            If SQL_Server Then
               sSQL = "UPDATE Asiento_RP_IE " _
                    & "SET " & CStr(.fields("Cod_Rol_Pago")) & " = CRR.Valor " _
                    & "FROM Asiento_RP_IE As ARP,Catalogo_Rol_Rubros As CRR "
            Else
               sSQL = "UPDATE Asiento_RP_IE As ARP,Catalogo_Rol_Rubros As CRR " _
                    & "SET ARP." & CStr(.fields("Cod_Rol_Pago")) & " = CRR.Valor "
            End If
            sSQL = sSQL _
                 & "WHERE CRR.Item = '" & NumEmpresa & "' " _
                 & "AND CRR.Periodo = '" & Periodo_Contable & "' " _
                 & "AND CRR.Cod_Rol_Pago = '" & CStr(.fields("Cod_Rol_Pago")) & "' " _
                 & "AND CRR.Mes = " & NoMes & " " _
                 & "AND ARP.Codigo = CRR.Codigo "
            Ejecutar_SQL_SP sSQL
           .MoveNext
         Loop
     End If
    End With
    
    If SQL_Server Then
       sSQL = "UPDATE Asiento_RP_IE " _
            & "SET SUELDO = CRR.Ing_Liquido " _
            & "FROM Asiento_RP_IE As ARP,Trans_Rol_Horas As CRR "
    Else
       sSQL = "UPDATE Asiento_RP_IE As ARP,Trans_Rol_Horas As CRR " _
            & "SET ARP.SUELDO = CRR.Ing_Liquido "
    End If
    sSQL = sSQL _
         & "WHERE CRR.Item = '" & NumEmpresa & "' " _
         & "AND CRR.Periodo = '" & Periodo_Contable & "' " _
         & "AND CRR.Fecha = #" & BuscarFecha(FechaRep) & "# " _
         & "AND CRR.Ing_Liquido > 0 " _
         & "AND ARP.Codigo = CRR.Codigo "
    Ejecutar_SQL_SP sSQL
    
    If SQL_Server Then
       sSQL = "UPDATE Asiento_RP_IE " _
            & "SET HORAS_EXT = CRR.Ing_Horas_Ext " _
            & "FROM Asiento_RP_IE As ARP,Trans_Rol_Horas As CRR "
    Else
       sSQL = "UPDATE Asiento_RP_IE As ARP,Trans_Rol_Horas As CRR " _
            & "SET ARP.HORAS_EXT = CRR.Ing_Horas_Ext "
    End If
    sSQL = sSQL _
         & "WHERE CRR.Item = '" & NumEmpresa & "' " _
         & "AND CRR.Periodo = '" & Periodo_Contable & "' " _
         & "AND CRR.Fecha = #" & BuscarFecha(FechaRep) & "# " _
         & "AND CRR.Horas_Exts > 0 " _
         & "AND ARP.Codigo = CRR.Codigo "
    Ejecutar_SQL_SP sSQL
    
    With AdoSRI.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            sSQL = "UPDATE Asiento_RP_IE " _
                 & "SET " & CStr(.fields("Cod_Rol_Pago")) & " = 0 " _
                 & "WHERE " & CStr(.fields("Cod_Rol_Pago")) & " IS NULL "
            Ejecutar_SQL_SP sSQL
           .MoveNext
         Loop
     End If
    End With
    sSQL = "UPDATE Asiento_RP_IE " _
         & "SET SUELDO = 0 " _
         & "WHERE SUELDO IS NULL "
    Ejecutar_SQL_SP sSQL
        
    sSQL = "UPDATE Asiento_RP_IE " _
         & "SET HORAS_EXT = 0 " _
         & "WHERE HORAS_EXT IS NULL "
    Ejecutar_SQL_SP sSQL
    
    sSQL = "SELECT * " _
         & "FROM Asiento_RP_IE " _
         & "ORDER BY Empleado "
    Select_Adodc_Grid DGIEEmpleados, AdoIEEmpleados, sSQL
End Sub

Public Sub Grabar_Ingresos_Egresos_Empleados()
Dim Sueldo As Currency
Dim Horas_Ext As Currency
    RatonReloj
    MiTiempo = Time
    DGIEEmpleados.Visible = False
    With AdoIEEmpleados.Recordset
     If .RecordCount > 0 Then
        .MoveLast
         CodigoB = .fields("I_E_Emp")
         sSQL = "DELETE * " _
              & "FROM Trans_Rol_Horas " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND Fecha = #" & BuscarFecha(FechaRep) & "# "
         Ejecutar_SQL_SP sSQL
         
         sSQL = "DELETE * " _
              & "FROM Catalogo_Rol_Rubros " _
              & "WHERE Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "AND Mes = " & NoMes & " " _
              & "AND I_E = '" & CodigoB & "' "
         Ejecutar_SQL_SP sSQL
         
        .MoveFirst
         Do While Not .EOF
            Codigo = .fields("Codigo")
            Sueldo = .fields("SUELDO")
            Horas_Ext = .fields("HORAS_EXT")
            If AdoSRI.Recordset.RecordCount > 0 Then
               AdoSRI.Recordset.MoveFirst
               Do While Not AdoSRI.Recordset.EOF
                  Cta = AdoSRI.Recordset.fields("Codigo")
                  Codigo1 = AdoSRI.Recordset.fields("Cod_Rol_Pago")
                  If IsNull(AdoSRI.Recordset.fields("Con_IESS")) Then
                     Bandera = False
                  Else
                     Bandera = AdoSRI.Recordset.fields("Con_IESS")
                  End If
                  Cuenta = AdoSRI.Recordset.fields("Cuenta")
                  Valor = .fields(Codigo1)
                  If Valor > 0 Then
                     SetAdoAddNew "Catalogo_Rol_Rubros"
                     SetAdoFields "I_E", CodigoB
                     SetAdoFields "Detalle", ULCase(Cuenta)
                     SetAdoFields "Mes", NoMes
                     SetAdoFields "Cta", Cta
                     SetAdoFields "TV", "V"
                     SetAdoFields "Valor", Valor
                     SetAdoFields "Codigo", Codigo
                     SetAdoFields "CPais", CodigoPais
                     SetAdoFields "Calc_IESS", Bandera
                     SetAdoFields "Cod_Rol_Pago", Codigo1
                     SetAdoFields "Periodo", Periodo_Contable
                     SetAdoFields "Item", NumEmpresa
                     SetAdoUpdate
                  End If
                  AdoSRI.Recordset.MoveNext
               Loop
            End If
            
            If (Sueldo + Horas_Ext) > 0 Then
               'MsgBox Sueldo & vbCrLf & Horas_Ext
               SetAdoAddNew "Trans_Rol_Horas"
               SetAdoFields "T", Val(adFalse)
               SetAdoFields "Dias", Day(FechaRep)
               SetAdoFields "Codigo", Codigo
               SetAdoFields "Fecha", FechaRep
               SetAdoFields "Horas", MiTiempo
               SetAdoFields "Valor_Hora", Redondear((Sueldo / Day(FechaRep)) / 8, 4)
               SetAdoFields "Horas_Exts", 1
               SetAdoFields "Ing_Liquido", Sueldo
               SetAdoFields "Ing_Horas_Ext", Horas_Ext
               SetAdoFields "CodigoU", CodigoUsuario
               SetAdoFields "Item", NumEmpresa
               SetAdoUpdate
            End If
           .MoveNext
         Loop
     End If
    End With
    RatonNormal
    DGIEEmpleados.Visible = True
    MsgBox "Proceso grabado exitosamente"
End Sub
