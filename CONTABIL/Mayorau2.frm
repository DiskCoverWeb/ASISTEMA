VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form MayorAux2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LISTADO DE PAGOS PENDIENTES POR PRESTAMOS"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "Mayorau2.frx":0000
      DataSource      =   "AdoCliente"
      Height          =   315
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Cliente"
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
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   105
      Top             =   5985
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataGridLib.DataGrid DGMayor 
      Bindings        =   "Mayorau2.frx":0019
      Height          =   4950
      Left            =   105
      TabIndex        =   3
      Top             =   945
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   8731
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
      Height          =   750
      Left            =   9555
      Picture         =   "Mayorau2.frx":0031
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   960
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
      Height          =   750
      Left            =   8505
      Picture         =   "Mayorau2.frx":08FB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   105
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7455
      Picture         =   "Mayorau2.frx":11C5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   210
      Top             =   1260
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
      CommandType     =   1
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
      Caption         =   "Cliente"
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
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Abonado"
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
      Left            =   5460
      TabIndex        =   11
      Top             =   5985
      Width           =   960
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CxC"
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
      Left            =   8085
      TabIndex        =   10
      Top             =   5985
      Width           =   750
   End
   Begin VB.Label LabelTotSaldo 
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
      Height          =   330
      Left            =   3780
      TabIndex        =   9
      Top             =   5985
      Width           =   1695
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Acreditado"
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
      Left            =   2625
      TabIndex        =   8
      Top             =   5985
      Width           =   1170
   End
   Begin VB.Label LabelTotDebe 
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
      Height          =   330
      Left            =   6405
      TabIndex        =   7
      Top             =   5985
      Width           =   1695
   End
   Begin VB.Label LabelTotHaber 
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
      Height          =   330
      Left            =   8820
      TabIndex        =   6
      Top             =   5985
      Width           =   1695
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   2010
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3545;582"
      Value           =   "0"
      Caption         =   "Buscar por Cliente:"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "MayorAux2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  sSQL = "SELECT C.Cliente,TP.T,TP.Fecha,TP.Credito_No,TP.Cuota_No,TP.Interes,TP.Capital,TP.Comision,TP.Pagos,TP.Saldo " _
       & "FROM Clientes As C,Trans_Prestamos As TP " _
       & "WHERE TP.Item = '" & NumEmpresa & "' "
  If CheckBox1.Value <> 0 Then
     Codigo = Ninguno
     With AdoCliente.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Cliente Like '" & DCCliente.Text & "' ")
          If Not .EOF Then Codigo = .Fields("Codigo")
      End If
     End With
     sSQL = sSQL & "AND TP.Cuenta_No = '" & Codigo & "' "
  End If
  sSQL = sSQL & "AND TP.Cuenta_No = C.Codigo " _
       & "ORDER BY C.Cliente,TP.Credito_No,TP.Cuota_No "
  SelectDataGrid DGMayor, AdoSubCta, sSQL, , True
  Debe = 0: Haber = 0
  With AdoSubCta.Recordset
   If .RecordCount > 0 Then
       Do While Not .EOF
          If .Fields("T") = "C" Then
              Debe = Debe + .Fields("Pagos")
          Else
              Haber = Haber + .Fields("Pagos")
          End If
         .MoveNext
       Loop
   End If
  End With
  LabelTotDebe.Caption = Format(Debe, "#,##0.00")
  LabelTotHaber.Caption = Format(Haber, "#,##0.00")
  LabelTotSaldo.Caption = Format(Debe + Haber, "#,##0.00")
End Sub

Private Sub Command2_Click()
  SQLMsg2 = ""
  SQLMsg1 = "LISTADO DE PAGOS PENDIENTES POR PRESTAMOS"
  If CheckBox1.Value <> 0 Then SQLMsg2 = DCCliente.Text
  ImprimirAdodc AdoSubCta, 1, 8, True
End Sub

Private Sub Command3_Click()
  Unload MayorAux2
End Sub

Private Sub Form_Activate()
 sSQL = "SELECT Cliente,Codigo " _
       & "FROM Clientes As C,Trans_Prestamos As TP " _
       & "WHERE TP.T = 'P' " _
       & "AND TP.Item = '" & NumEmpresa & "' " _
       & "AND TP.Cuenta_No = C.Codigo " _
       & "GROUP BY Cliente,Codigo "
  SelectDBCombo DCCliente, AdoCliente, sSQL, "Cliente"
 
  If Supervisor = False Then
     Command2.Enabled = CNivel_6
  End If
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm MayorAux2
  ConectarAdodc AdoSubCta
  ConectarAdodc AdoCliente
End Sub


