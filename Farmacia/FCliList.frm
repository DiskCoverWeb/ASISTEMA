VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FCliList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LISTAR CLIENTES"
   ClientHeight    =   6960
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FCliList.frx":0000
      Height          =   5055
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8916
      _Version        =   393216
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin MSAdodcLib.Adodc AdoCliente 
      Height          =   330
      Left            =   120
      Top             =   6600
      Width           =   10935
      _ExtentX        =   19288
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
      Caption         =   "LISTADO DE CLIENTES"
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
   Begin VB.CheckBox CheqMod 
      Caption         =   "Modificar"
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
      Left            =   9870
      TabIndex        =   7
      Top             =   1050
      Width           =   1275
   End
   Begin VB.TextBox TextPatron 
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
      Left            =   2730
      TabIndex        =   5
      Top             =   945
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.CheckBox CheqDir 
      Caption         =   "Ordenar por Dirección"
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
      Left            =   7455
      TabIndex        =   10
      Top             =   1050
      Width           =   2325
   End
   Begin VB.ListBox ListCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2730
      TabIndex        =   4
      Top             =   105
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   7455
      Picture         =   "FCliList.frx":0019
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   1170
   End
   Begin VB.CheckBox CheqBusqueda 
      Caption         =   "Con Patron de Busqueda"
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
      TabIndex        =   0
      Top             =   105
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   2535
      Begin VB.OptionButton OpcEjec 
         Caption         =   "&Ejecutivos"
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
         Top             =   525
         Width           =   1380
      End
      Begin VB.OptionButton OpcCli 
         Caption         =   "C&lientes"
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
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
   End
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
      Height          =   855
      Left            =   9975
      Picture         =   "FCliList.frx":045B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   105
      Width           =   1170
   End
   Begin VB.CommandButton CommandImp 
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
      Left            =   8715
      Picture         =   "FCliList.frx":0E51
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   105
      Width           =   1170
   End
End
Attribute VB_Name = "FCliList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheqBusqueda_Click()
  If CheqBusqueda.Value = 1 Then
     ListCliente.Visible = True
     TextPatron.Visible = True
  Else
     ListCliente.Visible = False
     TextPatron.Visible = False
  End If
End Sub

Private Sub Command1_Click()
  Unload FCliList
End Sub

Private Sub Command2_Click()
  TextoValido TextPatron, , True
  DGClientes.AllowUpdate = False
  If CheqMod.Value = 1 Then DGClientes.AllowUpdate = True
  sSQL = "SELECT Codigo,Cliente,Direccion,Telefono,Ciudad,Grupo,Empresa,FactM "
  sSQL = sSQL & "FROM Clientes "
  If OpcCli.Value Then
     sSQL = sSQL & "WHERE E = 'C' "
  Else
     sSQL = sSQL & "WHERE E <> 'C' "
  End If
  With AdoCliente.Recordset
    If CheqBusqueda.Value = 1 Then
       Select Case .Fields(ListCliente.Text).Type
         Case dbDate
              sSQL = sSQL & "AND " & ListCliente.Text & " = #" & BuscarFecha(TextPatron.Text) & "# "
         Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbBoolean
              sSQL = sSQL & "AND " & ListCliente.Text & " = " & Val(TextPatron.Text) & " "
         Case dbText, dbMemo
              LongStrg = Len(TextPatron.Text)
              sSQL = sSQL & "AND Mid(Ucase(" & ListCliente.Text & "),1," & LongStrg & ") = '" & TextPatron.Text & "' "
         Case Else
              LongStrg = Len(TextPatron.Text)
              sSQL = sSQL & "AND Mid(Ucase(" & ListCliente.Text & "),1," & LongStrg & ") = '" & TextPatron.Text & "' "
       End Select
    End If
  End With
  If CheqDir.Value = 1 Then
     sSQL = sSQL & "ORDER BY Direccion,Cliente "
  Else
     sSQL = sSQL & "ORDER BY Cliente "
  End If
  'SelectDataGrid DGClientes, AdoCliente, sSQL
End Sub

Private Sub CommandImp_Click()
  DGClientes.AllowUpdate = False
  CheqMod.Value = 0
  Mensajes = "Esta Seguro que desea Imprimir Clientes"
  Titulo = "Formulario de Impresion"
  If BoxMensaje = vbYes Then
     DGClientes.Visible = False
     MensajeEncabData = "L I S T A    D E    C L I E N T E S"
     ImprimirClientes AdoCliente, CheqDir.Value
     DGClientes.Visible = True
  End If
End Sub

Private Sub DGClientes_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then GenerarDataTexto FCliList, AdoCliente
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Codigo,Cliente,Empresa,Telefono,Ciudad,Direccion,Grupo,FactM "
  sSQL = sSQL & "FROM Clientes "
  sSQL = sSQL & "WHERE E = 'C' "
  sSQL = sSQL & "ORDER BY Cliente "
  'SelectDataGrid DGClientes, AdoCliente, sSQL
  ListCliente.Clear
  With AdoCliente.Recordset
      For I = 0 To .Fields.Count - 1
          ListCliente.AddItem .Fields(I).Name
      Next I
  End With
  ListCliente.Text = ListCliente.List(0)
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FCliList
  ConectarAdodc AdoCliente
End Sub

Private Sub ListCliente_DblClick()
  TextPatron.SetFocus
End Sub

Private Sub ListCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextPatron_GotFocus()
  MarcarTexto TextPatron
End Sub

Private Sub TextPatron_LostFocus()
  TextoValido TextPatron, , True
End Sub
