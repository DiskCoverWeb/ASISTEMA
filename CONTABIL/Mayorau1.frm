VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form MayorAux1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LISTADO DEL CATALOGO DE SUBCUENTAS"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   105
      Top             =   7035
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
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
      Caption         =   "SubCta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DGMayor 
      Bindings        =   "Mayorau1.frx":0000
      Height          =   6945
      Left            =   105
      TabIndex        =   8
      Top             =   105
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   12250
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
      Height          =   855
      Left            =   11025
      Picture         =   "Mayorau1.frx":0018
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3465
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Imprimir"
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
      Left            =   11025
      Picture         =   "Mayorau1.frx":08E2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   11025
      TabIndex        =   0
      Top             =   105
      Width           =   1170
      Begin VB.OptionButton OpcCC 
         Caption         =   "Pri&ma"
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
         TabIndex        =   9
         Top             =   1785
         Width           =   960
      End
      Begin VB.OptionButton OpcPM 
         Caption         =   "Pri&ma"
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
         TabIndex        =   5
         Top             =   1470
         Width           =   960
      End
      Begin VB.OptionButton OpcI 
         Caption         =   "I&ngreso"
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
         TabIndex        =   4
         Top             =   1155
         Width           =   960
      End
      Begin VB.OptionButton OpcG 
         Caption         =   "&Gastos"
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
         TabIndex        =   3
         Top             =   840
         Width           =   960
      End
      Begin VB.OptionButton OpcC 
         Caption         =   "Cx&C"
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
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OpcP 
         Caption         =   "Cx&P"
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
         Top             =   525
         Width           =   750
      End
   End
End
Attribute VB_Name = "MayorAux1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Insertar_SubCtas(MesNo As Byte, _
                            Mes As String, _
                            Cta As String, _
                            CodigoSubCta As String, _
                            TValor As Currency)
   SQL1 = "DELETE * " _
        & "FROM Trans_Presupuestos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Cta = '" & Cta & "' " _
        & "AND Codigo = '" & Codigo2 & "' " _
        & "AND Mes_No = " & MesNo & " "
   Ejecutar_SQL_SP SQL1
   SetAdoAddNew "Trans_Presupuestos"
   SetAdoFields "Mes_No", MesNo
   SetAdoFields "Cta", Cta
   SetAdoFields "Mes", UCaseStrg(MidStrg(Mes, 1, 3))
   SetAdoFields "Codigo", CodigoSubCta
   SetAdoFields "Presupuesto", TValor
   SetAdoUpdate
End Sub

Public Sub ListarSubCta(TipoSubCta As String)
  If OpcG.value Then
     DGMayor.Caption = "Catálogo - GASTOS"
  ElseIf OpcC.value Then
     DGMayor.Caption = "Catálogo - CUENTAS POR COBRAR"
  ElseIf OpcP.value Then
     DGMayor.Caption = "Catálogo - CUENTAS POR PAGAR"
  ElseIf OpcI.value Then
     DGMayor.Caption = "Catálogo - INGRESOS"
  ElseIf OpcPM.value Then
     DGMayor.Caption = "Catálogo - PRIMAS"
  Else
     DGMayor.Caption = "Catálogo - CENTRO DE COSTO"
  End If
  Select Case TipoSubCta
    Case "C", "P"
         sSQL = "SELECT C.Cliente,CC.Cuenta,CP.Codigo,CP.Cta,CP.TC " _
              & "FROM Catalogo_CxCxP As CP,Clientes As C,Catalogo_Cuentas As CC " _
              & "WHERE CP.TC = '" & TipoSubCta & "' " _
              & "AND CP.Item = '" & NumEmpresa & "' " _
              & "AND CP.Periodo = '" & Periodo_Contable & "' " _
              & "AND CP.Item = CC.Item " _
              & "AND CP.Periodo = CC.Periodo " _
              & "AND CP.Cta = CC.Codigo " _
              & "AND CP.Codigo = C.Codigo " _
              & "ORDER BY CP.Cta,C.Cliente "
    Case "CC", "I", "G", "PM"
         sSQL = "SELECT Detalle,Presupuesto,Codigo,TC " _
              & "FROM Catalogo_SubCtas " _
              & "WHERE TC = '" & TipoSubCta & "' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' " _
              & "ORDER BY Codigo,Detalle "
  End Select
  Select_Adodc_Grid DGMayor, AdoSubCta, sSQL
End Sub

Public Sub EliminarSubCta()
Dim TipoSubCta As String
  If OpcG.value Then
     TipoSubCta = "G"
  ElseIf OpcC.value Then
     TipoSubCta = "C"
  ElseIf OpcP.value Then
     TipoSubCta = "P"
  ElseIf OpcI.value Then
     TipoSubCta = "I"
  ElseIf OpcPM.value Then
     TipoSubCta = "PM"
  Else
     TipoSubCta = "CC"
  End If
  Select Case TipoSubCta
    Case "C", "P"
         sSQL = "DELETE * " _
              & "FROM Catalogo_CxCxP " _
              & "WHERE Cta = '" & Cta & "' " _
              & "AND Codigo = '" & Codigo & "' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' "
    Case "CC", "I", "G", "PM"
         sSQL = "DELETE * " _
              & "FROM Catalogo_SubCtas " _
              & "WHERE Codigo = '" & Codigo & "' " _
              & "AND Item = '" & NumEmpresa & "' " _
              & "AND Periodo = '" & Periodo_Contable & "' "
  End Select
  Ejecutar_SQL_SP sSQL
End Sub

Private Sub Command2_Click()
  SQLMsg1 = "":  SQLMsg2 = "":  SQLMsg3 = ""
  MensajeEncabData = DGMayor.Caption
  ImprimirAdodc AdoSubCta, 1, 8
End Sub

Private Sub Command3_Click()
  Unload MayorAux1
End Sub

Private Sub DGMayor_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then GenerarDataTexto MayorAux1, AdoSubCta
  If vbKeyDelete = KeyCode Then
     Codigo = DGMayor.Columns(0)
     Cadena = DGMayor.Columns(1)
     Cta = DGMayor.Columns(2)
     EliminarSubCta
     If OpcG.value Then
        ListarSubCta "G"
     ElseIf OpcC.value Then
        ListarSubCta "C"
     ElseIf OpcP.value Then
        ListarSubCta "P"
     ElseIf OpcI.value Then
        ListarSubCta "I"
     ElseIf OpcPM.value Then
        ListarSubCta "PM"
     Else
        ListarSubCta "CC"
     End If
  End If
End Sub

Private Sub Form_Activate()
  If Supervisor = False Then Command2.Enabled = CNivel(6)
  ListarSubCta "C"
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm MayorAux1
  ConectarAdodc AdoSubCta
End Sub

Private Sub OpcC_Click()
  ListarSubCta "C"
End Sub

Private Sub OpcCC_Click()
  ListarSubCta "CC"
End Sub

Private Sub OpcG_Click()
  ListarSubCta "G"
End Sub

Private Sub OpcI_Click()
  ListarSubCta "I"
End Sub

Private Sub OpcP_Click()
  ListarSubCta "P"
End Sub

Private Sub OpcPM_Click()
  ListarSubCta "PM"
End Sub

