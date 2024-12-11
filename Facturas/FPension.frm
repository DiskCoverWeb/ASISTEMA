VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FPensiones 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LISTAR CLIENTES POR GRUPO"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1588
      ButtonWidth     =   2328
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Insertar"
            Key             =   "Insertar"
            Object.ToolTipText     =   "Insertar Todos"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Eliminar"
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar Todos"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "(+/-) &Pension"
            Key             =   "Pension"
            Object.ToolTipText     =   "Cambia el Valor de la Pension"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "(+/-) &Desc."
            Key             =   "Descuento"
            Object.ToolTipText     =   "Cambia el Valor de Descuentos"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "(+/-) Desc. &2"
            Key             =   "Descuento2"
            Object.ToolTipText     =   "Cambia el Valor del Descuento 2"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Copiar Mes"
            Key             =   "Copiar_Mes"
            Object.ToolTipText     =   "Copia el mes en otros meses"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Multa"
            Key             =   "Multas"
            Object.ToolTipText     =   "Multas"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Recargos"
            Key             =   "s/Recargo"
            Object.ToolTipText     =   "Recargos"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrmCopiar 
      BackColor       =   &H00FF8080&
      Caption         =   "DUPLICAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   210
      TabIndex        =   15
      Top             =   2835
      Visible         =   0   'False
      Width           =   1800
      Begin VB.ListBox LstCopiar 
         BackColor       =   &H00C0FFFF&
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
         Height          =   2790
         Left            =   105
         TabIndex        =   16
         Top             =   210
         Width           =   1590
      End
   End
   Begin MSDataListLib.DataCombo DCInv 
      Bindings        =   "FPension.frx":0000
      DataSource      =   "AdoCxCxP"
      Height          =   4275
      Left            =   105
      TabIndex        =   14
      Top             =   2520
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7541
      _Version        =   393216
      Style           =   1
      BackColor       =   12632256
      ForeColor       =   8388608
      Text            =   "Productos"
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
   Begin MSDataListLib.DataCombo DCGrupoF 
      Bindings        =   "FPension.frx":0017
      DataSource      =   "AdoGrupo"
      Height          =   360
      Left            =   4095
      TabIndex        =   2
      Top             =   945
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   1065
      Left            =   105
      TabIndex        =   3
      Top             =   1365
      Width           =   11670
      Begin VB.TextBox TextCant 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6195
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "FPension.frx":002E
         Top             =   210
         Width           =   750
      End
      Begin VB.TextBox TxtArea 
         Alignment       =   1  'Right Justify
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
         Left            =   9975
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   210
         Width           =   1590
      End
      Begin VB.TextBox TxtDesc 
         Alignment       =   1  'Right Justify
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
         MaxLength       =   8
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   630
         Width           =   1485
      End
      Begin VB.TextBox TxtDesc2 
         Alignment       =   1  'Right Justify
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
         Left            =   6300
         MaxLength       =   8
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   630
         Width           =   1485
      End
      Begin MSMask.MaskEdBox MBFechaI 
         Height          =   330
         Left            =   2730
         TabIndex        =   5
         Top             =   210
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
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CANTIDAD DE MESES"
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
         Left            =   4095
         TabIndex        =   6
         Top             =   210
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA INICIO DE EMISION"
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
         TabIndex        =   4
         Top             =   210
         Width           =   2640
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VALOR A FACTURAR POR MES"
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
         Left            =   7035
         TabIndex        =   8
         Top             =   210
         Width           =   2955
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DESCUENTO POR MES"
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
         TabIndex        =   10
         Top             =   630
         Width           =   2220
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DESCUENTO 2 POR MES"
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
         Left            =   3885
         TabIndex        =   12
         Top             =   630
         Width           =   2430
      End
   End
   Begin MSDataListLib.DataCombo DCGrupoI 
      Bindings        =   "FPension.frx":0030
      DataSource      =   "AdoGrupo"
      Height          =   360
      Left            =   1890
      TabIndex        =   1
      Top             =   945
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox CheqRangos 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Por Rangos:"
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
      Top             =   945
      Value           =   1  'Checked
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
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
      Height          =   195
      Left            =   105
      Picture         =   "FPension.frx":0047
      TabIndex        =   17
      Top             =   5145
      Width           =   330
   End
   Begin MSAdodcLib.Adodc AdoCxCxP 
      Height          =   330
      Left            =   210
      Top             =   3360
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSAdodcLib.Adodc AdoRubros 
      Height          =   330
      Left            =   210
      Top             =   3045
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Rubros"
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
      Left            =   210
      Top             =   3675
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   3990
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2310
      Top             =   4725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FPension.frx":0911
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FPension.frx":0C2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FPension.frx":0F45
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FPension.frx":125F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FPension.frx":1579
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FPension.frx":1AEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FPension.frx":1E05
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FPensiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Contador = CInt(TextCant.Text)
    TextoValido TxtArea, True, True
    Valor = CCur(TxtArea.Text)
    Total_Desc = CCur(TxtDesc.Text)
    Total_Desc2 = CCur(TxtDesc2.Text)
    CodigoP = SinEspaciosIzq(DCInv.Text)
    Codigo1 = DCGrupoI.Text
    Codigo2 = DCGrupoF.Text
    FechaTexto = MBFechaI.Text
    Mifecha = MBFechaI.Text
    NoDiaT = Day(Mifecha)
    If Codigo1 = "" Then Codigo1 = Ninguno
    If Codigo2 = "" Then Codigo2 = Ninguno
   'MsgBox Button.key
    Select Case Button.key
      Case "Salir"
           Unload FPensiones
      Case "Insertar"
           Insertar_Pensiones
      Case "Eliminar"
           Eliminar_Pensiones
      Case "Pension"
           Tipo_Cambio_Valor "Pension"
      Case "Descuento"
           Tipo_Cambio_Valor "Descuento"
      Case "Descuento2"
           Tipo_Cambio_Valor "Descuento2"
      Case "Copiar_Mes"
           Copiar_Mes
      Case "Multas"
           Multas
      Case "Recargos"
           'Recargos
    End Select
    RatonNormal
End Sub

Private Sub CheqRangos_Click()
 If CheqRangos.value = 0 Then
    DCGrupoI.Enabled = False
    DCGrupoF.Enabled = False
 Else
    DCGrupoI.Enabled = True
    DCGrupoF.Enabled = True
 End If
End Sub

Public Sub Copiar_Mes()
 If ClaveSupervisor Then
    LstCopiar.Clear
    sSQL = "SELECT Periodo, Num_Mes " _
         & "FROM Clientes_Facturacion " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "GROUP BY Periodo, Num_Mes " _
         & "ORDER BY Periodo, Num_Mes "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            LstCopiar.AddItem .fields("Periodo") & " " & Format$(.fields("Num_Mes"), "00")
           .MoveNext
         Loop
         FrmCopiar.Visible = True
         'MsgBox FechaTexto
         LstCopiar.Text = LstCopiar.List(0)
        'MBFechaI.Text = FechaTexto
         LstCopiar.SetFocus
     Else
       MsgBox "No existe datos para copiar"
       FrmCopiar.Visible = False
     End If
    End With
 End If
End Sub

Private Sub Command2_Click()
  Unload FPensiones
End Sub

Public Sub Eliminar_Pensiones()
    If ClaveSupervisor And Contador >= 1 Then
       RatonReloj
       FPensiones.Caption = Format$(0 / Contador, "00%") & "ELIMINANDO DATOS ANTERIORES"
       
       Rago_Fechas_Proceso
       
       sSQL = "DELETE * " _
            & "FROM Clientes_Facturacion " _
            & "WHERE Codigo_Inv = '" & CodigoP & "' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Fecha BETWEEN '" & BuscarFecha(FechaTexto) & "' and '" & BuscarFecha(Mifecha) & "' "
       If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
       Ejecutar_SQL_SP sSQL
       RatonNormal
       MsgBox "Proceso Terminado"
       Unload FPensiones
    Else
       RatonNormal
       MsgBox "No se puede realizar este proceso"
    End If
End Sub

Public Sub Insertar_Pensiones()
Dim Proceder As Boolean
    If ClaveSupervisor And Contador >= 1 Then
       RatonReloj
       Proceder = True
       FPensiones.Caption = Format$(0 / Contador, "00%") & "ELIMINANDO DATOS ANTERIORES"
       
       Rago_Fechas_Proceso
       
       sSQL = "SELECT TOP 10 Codigo " _
            & "FROM Clientes_Facturacion " _
            & "WHERE Codigo_Inv = '" & CodigoP & "' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND Fecha BETWEEN '" & BuscarFecha(FechaTexto) & "' and '" & BuscarFecha(Mifecha) & "' "
       If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
       Select_Adodc AdoAux, sSQL
       
       If AdoAux.Recordset.RecordCount > 0 Then
          Titulo = "GENERACION DE RUBROS A FACTURAR POR LOTE "
          Mensajes = "Actualmente ya existe rubros a facturar en este rango de Grupo y de fechas." & vbCrLf & vbCrLf _
                   & "Realmente desea borrar estos datos he ingresar los nuevos "
          Mensajes = UCaseStrg(Mensajes)
          If BoxMensaje <> vbYes Then Proceder = False
       End If
       If Proceder Then
          sSQL = "DELETE * " _
               & "FROM Clientes_Facturacion " _
               & "WHERE Codigo_Inv = '" & CodigoP & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Fecha BETWEEN '" & BuscarFecha(FechaTexto) & "' and '" & BuscarFecha(Mifecha) & "' "
          If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
          Ejecutar_SQL_SP sSQL
          Mifecha = MBFechaI.Text
          For I = 1 To Contador
              NoDias = Day(Mifecha)
              NoMes = Month(Mifecha)
              NoAnio = Year(Mifecha)
              Mes1 = MesesLetras(NoMes)
              FPensiones.Caption = Format$(I / Contador, "00%") & " - ASIGNACION DE CODIGOS DE FACTURARION - GRUPO: " & Codigo1 & " - " & Codigo2
              sSQL = "INSERT INTO Clientes_Facturacion (T, GrupoNo, Codigo, Codigo_Inv, Valor, Descuento, Descuento2, CodigoU, Item, Periodo, Num_Mes, Mes, Fecha) " _
                   & "SELECT 'N', Grupo, Codigo, '" & CodigoP & "', " & CStr(Valor) & ", " & CStr(Total_Desc) & ", " & CStr(Total_Desc2) & ", '" & CodigoUsuario & "', '" & NumEmpresa & "', '" & NoAnio & "', " _
                   & NoMes & ", '" & Mes1 & "', '" & BuscarFecha(Mifecha) & "' " _
                   & "FROM Clientes " _
                   & "WHERE FA <> " & Val(adFalse) & " "
              If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
              If CheqRangos.value <> 0 Then sSQL = sSQL & "AND Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
              sSQL = sSQL & "ORDER BY Grupo, Cliente, Sexo "
              Ejecutar_SQL_SP sSQL
             'MsgBox I & ":" & vbCrLf & sSQL
              Mifecha = CLongFecha(CFechaLong(Mifecha) + 31)
              NoDias = Day(Mifecha)
              NoMes = Month(Mifecha)
              NoAnio = Year(Mifecha)
              Mifecha = Format(NoDiaT, "00") & "/" & Format(NoMes, "00") & "/" & Format(NoAnio, "0000")
          Next I
          Eliminar_Nulos_SP "Clientes_Facturacion"
          RatonNormal
          MsgBox "Proceso Terminado"
          Unload FPensiones
       Else
          RatonNormal
          MsgBox "No se realizo ningun proceso"
       End If
    End If
End Sub

Public Sub Tipo_Cambio_Valor(Tipo_Cambio As String)
Dim Valor_Cambiar As Currency
     Valor_Cambiar = 0
     Select Case Tipo_Cambio
       Case "Pension":    Valor_Cambiar = Val(TxtArea)
       Case "Descuento":  Valor_Cambiar = Val(TxtDesc)
       Case "Descuento2": Valor_Cambiar = Val(TxtDesc2)
     End Select
     If Valor_Cambiar <> 0 And Contador >= 1 Then
        Titulo = "CAMBIO DE VALORES EN GRUPO"
        Select Case Tipo_Cambio
          Case "Pension":    Mensajes = "AUMENTA/DECREMENTA LOS VALORES DE PENSION: " & vbCrLf & vbCrLf
          Case "Descuento":  Mensajes = "AUMENTA/DECREMENTA LOS VALORES DE DESCUENTOS: " & vbCrLf & vbCrLf
          Case "Descuento2": Mensajes = "AUMENTA/DECREMENTA LOS VALORES DE DESCUENTOS2: " & vbCrLf & vbCrLf
        End Select
        Mensajes = Mensajes _
                 & UCaseStrg(DCInv.Text) & vbCrLf & vbCrLf _
                 & "POR: " & Format(Valor_Cambiar, "#0.00")
        If BoxMensaje = vbYes Then
           Rago_Fechas_Proceso
           
           sSQL = "UPDATE Clientes_Facturacion "
           Select Case Tipo_Cambio
             Case "Pension":    sSQL = sSQL & "SET Valor = Valor "
             Case "Descuento":  sSQL = sSQL & "SET Descuento = Descuento "
             Case "Descuento2": sSQL = sSQL & "SET Descuento2 = Descuento2 "
           End Select
           If Valor_Cambiar > 0 Then sSQL = sSQL & "+" & Valor_Cambiar & " " Else sSQL = sSQL & Valor_Cambiar & " "
           
           sSQL = sSQL & "WHERE Codigo_Inv = '" & CodigoP & "' " _
                & "AND Item = '" & NumEmpresa & "' " _
                & "AND Fecha BETWEEN '" & BuscarFecha(FechaTexto) & "' and '" & BuscarFecha(Mifecha) & "' "
           If Tipo_Cambio = "Descuento2" Then sSQL = sSQL & "AND Descuento = 0 "
           If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
           
           'MsgBox sSQL
           
           Ejecutar_SQL_SP sSQL
           Select Case Tipo_Cambio
             Case "Pension"
                  sSQL = "DELETE * " _
                       & "FROM Clientes_Facturacion " _
                       & "WHERE Valor <= 0 "
             Case "Descuento"
                  sSQL = "UPDATE Clientes_Facturacion " _
                       & "SET Descuento = 0 " _
                       & "WHERE Descuento < 0 "
             Case "Descuento2"
                  sSQL = "UPDATE Clientes_Facturacion " _
                       & "SET Descuento2 = 0 " _
                       & "WHERE Descuento2 < 0 "
           End Select
           sSQL = sSQL & "AND Item = '" & NumEmpresa & "' "
           Ejecutar_SQL_SP sSQL
           MsgBox "Proceso Terminado"
           Unload FPensiones
        End If
     End If
End Sub

Public Sub Multas()
  If ClaveSupervisor And Contador >= 1 Then
     Actualizar_Abonos_Facturas_SP FA, True
     Rago_Fechas_Proceso
     
     sSQL = "DELETE * " _
          & "FROM Clientes_Facturacion " _
          & "WHERE Codigo_Inv = '" & CodigoP & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Fecha BETWEEN '" & BuscarFecha(FechaTexto) & "' and '" & BuscarFecha(Mifecha) & "' "
     If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
     Ejecutar_SQL_SP sSQL

     sSQL = "UPDATE Clientes " _
          & "SET X = 'M' " _
          & "WHERE FA <> " & Val(adFalse) & " "
     Ejecutar_SQL_SP sSQL
     
     sSQL = "UPDATE Clientes " _
          & "SET X = 'F' " _
          & "FROM Clientes As C, Facturas As F " _
          & "WHERE F.Item = '" & NumEmpresa & "' " _
          & "AND F.TC IN ('FA','FM','NV') " _
          & "AND F.T <> 'A' " _
          & "AND F.Fecha BETWEEN #" & BuscarFecha(FechaTexto) & "# and #" & BuscarFecha(Mifecha) & "# "
     If CheqRangos.value <> 0 Then sSQL = sSQL & "AND C.Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
     sSQL = sSQL & "AND  C.Codigo = F.CodigoC  "
     Ejecutar_SQL_SP sSQL
      
     Mifecha = MBFechaI.Text
     For I = 1 To Contador
         NoDias = Day(Mifecha)
         NoMes = Month(Mifecha)
         NoAnio = Year(Mifecha)
         Mes1 = MesesLetras(NoMes)
         FPensiones.Caption = Format$(I / Contador, "00%") & " - ASIGNACION DE CODIGOS DE FACTURARION - GRUPO: " & Codigo1 & " - " & Codigo2
         sSQL = "INSERT INTO Clientes_Facturacion (T, GrupoNo, Codigo, Codigo_Inv, Valor, CodigoU, Item, Periodo, Num_Mes, Mes, Fecha) " _
              & "SELECT 'N', Grupo, Codigo, '" & CodigoP & "', " & CStr(Valor) & ", '" & CodigoUsuario & "', '" & NumEmpresa & "', '" & NoAnio & "', " _
              & NoMes & ", '" & Mes1 & "', '" & BuscarFecha(Mifecha) & "' " _
              & "FROM Clientes " _
              & "WHERE FA <> " & Val(adFalse) & " " _
              & "AND X = 'M' "
         If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
         If CheqRangos.value <> 0 Then sSQL = sSQL & "AND Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
         sSQL = sSQL & "ORDER BY Grupo, Cliente, Sexo "
         Ejecutar_SQL_SP sSQL
         'MsgBox I & ":" & vbCrLf & sSQL
         Mifecha = CLongFecha(CFechaLong(Mifecha) + 31)
         NoDias = Day(Mifecha)
         NoMes = Month(Mifecha)
         NoAnio = Year(Mifecha)
         Mifecha = Format(NoDiaT, "00") & "/" & Format(NoMes, "00") & "/" & Format(NoAnio, "0000")
     Next I
     Eliminar_Nulos_SP "Clientes_Facturacion"
     RatonNormal
     MsgBox "Proceso Terminado"
     Unload FPensiones
  Else
     RatonNormal
     MsgBox "No se puede proceder"
     Unload FPensiones
  End If
End Sub

'''Public Sub Recargos()
''''RECARGO
'''  If ClaveSupervisor Then
'''    If LstMeses.ListCount > 1 Then
'''       TextoValido TxtArea, True, True
'''       Valor = CCur(TxtArea.Text)
'''       CodigoP = SinEspaciosIzq(DCInv.Text)
'''       Codigo1 = DCGrupoI.Text
'''       Codigo2 = DCGrupoF.Text
'''       If Codigo1 = "" Then Codigo1 = Ninguno
'''       If Codigo2 = "" Then Codigo2 = Ninguno
'''       If Valor > 0 Then
'''          For NoMes = 1 To 12
'''           FPensiones.Caption = Format$(NoMes / 12, "00%") & " - ASIGNACION DE RECARGOS "
'''           If LstMeses.Selected(NoMes) Then
'''              FechaInicial = "01/" & Format$(NoMes, "00") & "/" & Format$(Codigo4, "0000")
'''              FechaFinal = UltimoDiaMes(FechaInicial)
'''              FechaIni = BuscarFecha(FechaInicial)
'''              FechaFin = BuscarFecha(FechaFinal)
'''              Contador = 0
'''              sSQL = "SELECT C.Grupo,C.Codigo,C.Cliente,SUM(CF.Valor) As Valor_Total,COUNT(C.Codigo) As Cantidad " _
'''                   & "FROM Clientes As C, Clientes_Facturacion As CF " _
'''                   & "WHERE CF.Num_Mes = " & NoMes & " " _
'''                   & "AND CF.Periodo = '" & Codigo4 & "' " _
'''                   & "AND CF.Codigo_Inv <> '" & CodigoP & "' " _
'''                   & "AND CF.Item = '" & NumEmpresa & "' "
'''              If CheqRangos.value <> 0 Then
'''                 sSQL = sSQL & "AND C.Grupo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
'''              End If
'''              sSQL = sSQL _
'''                   & "AND C.Codigo = CF.Codigo " _
'''                   & "GROUP BY C.Grupo,C.Codigo,C.Cliente " _
'''                   & "ORDER BY C.Grupo,C.Codigo,C.Cliente "
'''             'MsgBox sSQL
'''              Select_Adodc AdoRubros, sSQL
'''              With AdoRubros.Recordset
'''               If .RecordCount > 0 Then
'''                   Do While Not .EOF
'''                      Contador = Contador + 1
'''                      CodigoCliente = .fields("Codigo")
'''                      Codigo1 = .fields("Grupo")
'''                      NombreCliente = .fields("Cliente")
'''                      sSQL = "SELECT * " _
'''                           & "FROM Clientes_Facturacion " _
'''                           & "WHERE Codigo_Inv = '" & CodigoP & "' " _
'''                           & "AND Codigo = '" & CodigoCliente & "' " _
'''                           & "AND Num_Mes = " & NoMes & " " _
'''                           & "AND Periodo = '" & Codigo4 & "' " _
'''                           & "AND Item = '" & NumEmpresa & "' "
'''                      Select_Adodc AdoAux, sSQL
'''                      If AdoAux.Recordset.RecordCount <= 0 Then
'''                         FPensiones.Caption = Format$(NoMes / 12, "00%") & " - INSERTAR A: " & NombreCliente & " EL RECARGO"
'''                         SetAdoAddNew "Clientes_Facturacion"
'''                         SetAdoFields "T", Normal
'''                         SetAdoFields "Codigo", CodigoCliente
'''                         SetAdoFields "Valor", Valor
'''                         SetAdoFields "Codigo_Inv", CodigoP
'''                         SetAdoFields "Num_Mes", NoMes
'''                         SetAdoFields "GrupoNo", Codigo1
'''                         SetAdoFields "Mes", MesesLetras(CInt(NoMes))
'''                         SetAdoFields "Item", NumEmpresa
'''                         SetAdoFields "Periodo", Codigo4
'''                         SetAdoFields "Fecha", FechaFinal
'''                         SetAdoFields "CodigoU", CodigoUsuario
'''                         SetAdoUpdate
'''                      End If
'''                     .MoveNext
'''                   Loop
'''               End If
'''              End With
'''           End If
'''          Next NoMes
'''       Else
'''           MsgBox "No se puede poner recargo en cero"
'''       End If
'''    End If
'''  End If
'''  MsgBox "Fin del proceso, Verfique los resultados"
'''  Unload FPensiones
'''End Sub

Private Sub Form_Activate()
  FPensiones.Caption = "CUENTAS POR COBRAR EXTRACONTABLE"
  sSQL = "SELECT Codigo_Inv & '  ' & Producto As NomProd,* " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "AND LEN(Cta_Inventario) = 1 " _
       & "AND INV <> " & Val(adFalse) & " " _
       & "ORDER BY Codigo_Inv "
  SelectDB_Combo DCInv, AdoCxCxP, sSQL, "NomProd"
  
  sSQL = "SELECT Grupo " _
       & "FROM Clientes " _
       & "WHERE FA <> " & Val(adFalse) & " "
  If Mas_Grupos Then sSQL = sSQL & "AND DirNumero = '" & NumEmpresa & "' "
  sSQL = sSQL _
       & "AND Cliente <> 'CONSUMIDOR FINAL' " _
       & "GROUP BY Grupo " _
       & "ORDER BY Grupo "
  SelectDB_Combo DCGrupoI, AdoGrupo, sSQL, "Grupo"
  SelectDB_Combo DCGrupoF, AdoGrupo, sSQL, "Grupo", True
''  If AdoGrupo.Recordset.RecordCount > 0 Then
''     AdoGrupo.Recordset.MoveLast
''     DCGrupoF.Text = AdoGrupo.Recordset.fields("Grupo")
''  End If
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FPensiones
  ConectarAdodc AdoAux
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoCxCxP
  ConectarAdodc AdoRubros
  FPensiones.Caption = "ASIGNACION DE CODIGO DE FACTURACION   GRUPO: " & Codigo1
  TxtArea = "0.00"
  TxtDesc = "0.00"
  TxtDesc2 = "0.00"
  MBFechaI.Text = FechaSistema
  TextCant = "0"
End Sub

Private Sub LstCopiar_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Copiar_Periodo As String
Dim Copiar_Mes As Integer
Dim FechaSigMes As String
  If KeyCode = vbKeyReturn Then
     If Contador >= 1 Then
        Copiar_Periodo = SinEspaciosIzq(LstCopiar.Text)
        Copiar_Mes = Val(SinEspaciosDer(LstCopiar.Text))
        Rago_Fechas_Proceso
        sSQL = "DELETE * " _
             & "FROM Clientes_Facturacion " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Fecha BETWEEN '" & BuscarFecha(FechaTexto) & "' and '" & BuscarFecha(Mifecha) & "' "
        If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
        Ejecutar_SQL_SP sSQL
        'MsgBox sSQL
        Mifecha = MBFechaI.Text
        For I = 1 To Contador
            NoDias = Day(Mifecha)
            NoMes = Month(Mifecha)
            NoAnio = Year(Mifecha)
            Mes1 = MesesLetras(NoMes)
            FPensiones.Caption = Format$(I / Contador, "00%") & " - ASIGNACION DE CODIGOS DE FACTURARION - GRUPO: " & Codigo1 & " - " & Codigo2
            sSQL = "INSERT INTO Clientes_Facturacion (T, GrupoNo, Codigo, Codigo_Inv, Valor, Item, CodigoU, Periodo, Num_Mes, Mes, Fecha, D) " _
                 & "SELECT 'N', GrupoNo, Codigo, Codigo_Inv, Valor, Item, '" & CodigoUsuario & "', '" & NoAnio & "', " _
                 & NoMes & ", '" & Mes1 & "', '" & BuscarFecha(Mifecha) & "', 0 " _
                 & "FROM Clientes_Facturacion " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Copiar_Periodo & "' " _
                 & "AND Num_Mes = " & Copiar_Mes & " "
            If CheqRangos.value <> 0 Then sSQL = sSQL & "AND GrupoNo BETWEEN '" & Codigo1 & "' and '" & Codigo2 & "' "
            sSQL = sSQL & "ORDER BY GrupoNo, Codigo, Codigo_Inv "
           'MsgBox sSQL
            Ejecutar_SQL_SP sSQL
           'MsgBox I & ":" & vbCrLf & sSQL
            Mifecha = CLongFecha(CFechaLong(Mifecha) + 31)
            NoDias = Day(Mifecha)
            NoMes = Month(Mifecha)
            NoAnio = Year(Mifecha)
            Mifecha = Format(NoDiaT, "00") & "/" & Format(NoMes, "00") & "/" & Format(NoAnio, "0000")
        Next I
        Eliminar_Nulos_SP "Clientes_Facturacion"
        RatonNormal
        MsgBox "Proceso Terminado"
        Unload FPensiones
     End If
  End If
  If KeyCode = vbKeyEscape Then
     FrmCopiar.Visible = False
     'Command1.SetFocus
  End If
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

Private Sub TextCant_GotFocus()
    MarcarTexto TextCant
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TextCant_LostFocus()
    TextoValido TextCant, True, , 0
End Sub

Private Sub TxtArea_GotFocus()
    MarcarTexto TxtArea
End Sub

Private Sub TxtArea_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtArea_LostFocus()
  TextoValido TxtArea, True, , 2
End Sub

Private Sub TxtDesc_GotFocus()
  MarcarTexto TxtDesc
End Sub

Private Sub TxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDesc_LostFocus()
  TextoValido TxtDesc, True, , 2
End Sub

Private Sub TxtDesc2_GotFocus()
  MarcarTexto TxtDesc2
End Sub

Private Sub TxtDesc2_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtDesc2_LostFocus()
  TextoValido TxtDesc2, True, , 2
End Sub

Public Sub Rago_Fechas_Proceso()
    For I = 1 To Contador
        NoDias = Day(Mifecha)
        NoMes = Month(Mifecha)
        NoAnio = Year(Mifecha)
        Mes1 = MesesLetras(NoMes)
        If I < Contador Then
           Mifecha = CLongFecha(CFechaLong(Mifecha) + 31)
           NoDias = Day(Mifecha)
           NoMes = Month(Mifecha)
           NoAnio = Year(Mifecha)
           Mifecha = Format(NoDiaT, "00") & "/" & Format(NoMes, "00") & "/" & Format(NoAnio, "0000")
        End If
    Next I
    Mifecha = UltimoDiaMes(Mifecha)
End Sub
