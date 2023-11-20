VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FCodigosRetencion 
   Caption         =   "CATALOGO DE CODIGOS DE RETENCIONES POR OTROS CONCEPTOS"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   7485
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CPeriodo 
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
      Left            =   2205
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   105
      Width           =   3900
   End
   Begin MSAdodcLib.Adodc AdoSubCta 
      Height          =   330
      Left            =   105
      Top             =   5985
      Visible         =   0   'False
      Width           =   3060
      _ExtentX        =   5398
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
   Begin MSDataGridLib.DataGrid DGSubCta 
      Bindings        =   "FCodigos.frx":0000
      Height          =   5580
      Left            =   105
      TabIndex        =   2
      Top             =   420
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   9843
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
      Left            =   6195
      Picture         =   "FCodigos.frx":0018
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1050
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
      Left            =   6195
      Picture         =   "FCodigos.frx":08E2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1170
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Periodo de Retención:"
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
      Width           =   2115
   End
End
Attribute VB_Name = "FCodigosRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
  SQLMsg1 = "": SQLMsg2 = "": SQLMsg3 = ""
  MensajeEncabData = "CATALOGO DE CODIGOS DE RETENCION"
  Imprimir_Catalogo_Ret AdoSubCta, 1, 7, True
End Sub

Private Sub Command3_Click()
  Unload FCodigosRetencion
End Sub

Private Sub CPeriodo_GotFocus()
  'MarcarTexto
End Sub

Private Sub CPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub CPeriodo_LostFocus()
  sSQL = "SELECT Codigo,Concepto,(Porcentaje*1) As Porcentejes,Ingresar_Porcentaje,Fecha_Inicio,Fecha_Final,T " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' "
  If CPeriodo.Text <> "TODOS" Then sSQL = sSQL & "AND Fecha_Inicio = #" & BuscarFecha(CPeriodo.Text) & "# "
  sSQL = sSQL & "ORDER BY Codigo,Fecha_Inicio,Fecha_Final "
  Select_Adodc_Grid DGSubCta, AdoSubCta, sSQL
End Sub

Private Sub DGSubCta_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then GenerarDataTexto FCodigosRetencion, AdoSubCta
End Sub

Private Sub Form_Activate()
  DGSubCta.width = MDI_X_Max - 1800
  DGSubCta.Height = MDI_Y_Max - 400
  DGSubCta.Caption = "CATALOGO DE CODIGOS DE RETENCION POR OTROS CONCEPTOS"
  Command2.Left = MDI_X_Max - 1600
  Command3.Left = MDI_X_Max - 1600
  CPeriodo.Clear
  sSQL = "SELECT Fecha_Inicio " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "GROUP BY Fecha_Inicio " _
       & "ORDER BY Fecha_Inicio DESC "
  Select_Adodc AdoSubCta, sSQL
  With AdoSubCta.Recordset
   If .RecordCount Then
       Do While Not .EOF
          CPeriodo.AddItem .Fields("Fecha_Inicio")
         .MoveNext
       Loop
   End If
  End With
  CPeriodo.AddItem "TODOS"
  CPeriodo.Text = CPeriodo.List(0)
  sSQL = "SELECT Codigo,Concepto,(Porcentaje*1) As Porcentejes,Ingresar_Porcentaje,Fecha_Inicio,Fecha_Final,T " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' "
  If CPeriodo.Text <> "TODOS" Then sSQL = sSQL & "AND Fecha_Inicio = #" & BuscarFecha(CPeriodo.Text) & "# "
  sSQL = sSQL & "ORDER BY Codigo,Fecha_Inicio,Fecha_Final "
  Select_Adodc_Grid DGSubCta, AdoSubCta, sSQL
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoSubCta
End Sub

