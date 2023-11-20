VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FGarantes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LISTADO DE GARANTES"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
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
      Left            =   10710
      TabIndex        =   0
      Top             =   6825
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DGGarantes 
      Bindings        =   "FGarante.frx":0000
      Height          =   6735
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   11880
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
   Begin MSAdodcLib.Adodc AdoGarantes 
      Height          =   330
      Left            =   105
      Top             =   6825
      Width           =   10620
      _ExtentX        =   18733
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
      Caption         =   "Saldos"
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
Attribute VB_Name = "FGarantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload FGarantes
End Sub

Private Sub DGGarantes_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGGarantes.Visible = False
     GenerarDataTexto FGarantes, AdoGarantes, True
     DGGarantes.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyB Then BuscarDatos DGGarantes, AdoGarantes
  If CtrlDown And KeyCode = vbKeyP Then
     DGGarantes.Visible = False
     SQLMsg1 = "REPORTE DE GARANTES"
     ImprimirAdodc AdoGarantes, 2, 8
     DGGarantes.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyF5 Then
     If ClaveContador Then
        DGGarantes.AllowUpdate = True
        DGGarantes.AllowDelete = True
        MsgBox "Proceso aceptado, puede empezar a modificar o eliminar"
     End If
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT TP,Cuenta_No,Credito_No,Nombres,CI,Telefono,LugarTrabajo,Direccion  " _
       & "FROM Clientes_Datos_Extras " _
       & "WHERE Cuenta_No <> '.' " _
       & "ORDER BY Cuenta_No,Credito_No,GC,Nombres "
  SQLDec = ""
  SelectDataGrid DGGarantes, AdoGarantes, sSQL, SQLDec
End Sub

Private Sub Form_Load()
  CentrarForm FGarantes
  ConectarAdodc AdoGarantes
End Sub
