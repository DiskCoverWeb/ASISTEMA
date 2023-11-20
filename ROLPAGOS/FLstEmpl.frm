VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FListarEmpleados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LISTA DE EMPLEADOS"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   855
   End
   Begin MSAdodcLib.Adodc AdoListRol 
      Height          =   330
      Left            =   105
      Top             =   6510
      Width           =   3480
      _ExtentX        =   6138
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
   Begin MSDataGridLib.DataGrid DGListRol 
      Bindings        =   "FLstEmpl.frx":0000
      Height          =   6000
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   10583
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
End
Attribute VB_Name = "FListarEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub DGListRol_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
     Codigo = AdoListRol.Recordset.Fields("Codigo")
     NombreCliente = AdoListRol.Recordset.Fields("Empleado")
     Titulo = "Pregunta de Eliminacion"
     Mensajes = "Eliminar: " & NombreCliente
     If BoxMensaje = vbYes Then
        SQL2 = "DELETE * " _
             & "FROM Catalogo_RolPagos " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Codigo = '" & Codigo & "' "
        ConectarAdoExecute SQL2
        sSQL = "SELECT CR.Codigo,C.Cliente As Empleado,C.CI_RUC,CR.Grupo_Rol " _
             & "FROM Clientes As C, Catalogo_RolPagos As CR " _
             & "WHERE CR.Item = '" & NumEmpresa & "' " _
             & "AND C.Codigo = CR.Codigo " _
             & "ORDER BY C.Cliente "
        SelectDataGrid DGListRol, AdoListRol, sSQL
        Command1.SetFocus
     End If
  End If
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT CR.Codigo,C.Cliente As Empleado,C.CI_RUC,CR.Grupo_Rol " _
       & "FROM Clientes As C, Catalogo_RolPagos As CR " _
       & "WHERE CR.Item = '" & NumEmpresa & "' " _
       & "AND C.Codigo = CR.Codigo " _
       & "ORDER BY C.Cliente "
  SelectDataGrid DGListRol, AdoListRol, sSQL
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FListarEmpleados
  ConectarAdodc AdoListRol
End Sub
