VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form CatalogoCtas 
   Caption         =   "PRESENTACION DEL CATALOGO DE CUENTA"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CheqPM 
      Caption         =   "Solo Productos de Movimiento"
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
      Left            =   3675
      TabIndex        =   8
      Top             =   105
      Width           =   4320
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "Catalogo.frx":0000
      Height          =   5895
      Left            =   105
      TabIndex        =   7
      ToolTipText     =   "<F1> Genera Archivo,  <Ctrl>+<F5> Modificar datos"
      Top             =   945
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   10398
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   105
      Top             =   6825
      Width           =   11460
      _ExtentX        =   20214
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
   Begin VB.CommandButton Command3 
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
      Left            =   8190
      Picture         =   "Catalogo.frx":0017
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10500
      Picture         =   "Catalogo.frx":0459
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9345
      Picture         =   "Catalogo.frx":0D23
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   1065
   End
   Begin MSMask.MaskEdBox MBoxCtaF 
      Height          =   330
      Left            =   1890
      TabIndex        =   3
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MBoxCtaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta Inicial"
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
      Width           =   1695
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cuenta Final"
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
      Left            =   1890
      TabIndex        =   2
      Top             =   105
      Width           =   1695
   End
End
Attribute VB_Name = "CatalogoCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''
'''Private Sub CopySelectedPictureToClipboard(myFlex As MSHFlexGrid)
'''   Dim I As Integer, tr As Long, lc As Long, hl As Integer
'''   ' Se prepara para funcionar.
'''   myFlex.Redraw = False  ' Para eliminar el parpadeo.
'''   hl = myFlex.HighLight  ' Guarda los valores actuales.
'''   tr = myFlex.TopRow
'''   lc = myFlex.LeftCol
'''   myFlex.HighLight = 0  ' No se resalta la imagen.
'''   ' Oculta las filas y columnas no seleccionadas.
'''   ' (Guarda los tamaños originales en las propiedades
'''   ' RowData y ColData).
'''   For I = myFlex.FixedRows To myFlex.Rows - 1
'''      If I < myFlex.Row Or I > myFlex.RowSel Then
'''         myFlex.RowData(I) = myFlex.RowHeight(I)
'''         myFlex.RowHeight(I) = 0
'''      End If
'''   Next
'''   For I = myFlex.FixedCols To myFlex.Cols - 1
'''      If I < myFlex.Col Or I > myFlex.ColSel Then
'''         myFlex.ColData(I) = myFlex.ColWidth(I)
'''         myFlex.ColWidth(I) = 0
'''      End If
'''   Next
'''   ' Se desplaza a la esquina superior izquierda.
'''   myFlex.TopRow = myFlex.FixedRows
'''   myFlex.LeftCol = myFlex.FixedCols
'''   ' Copia la imagen.
'''   Clipboard.Clear
'''   On Error Resume Next
'''   myFlex.PictureType = 0 ' Color.
'''   Clipboard.SetData myFlex.Picture
'''   If Error <> 0 Then
'''      myFlex.PictureType = 1 ' Monocromo.
'''      Clipboard.SetData myFlex.Picture
'''   End If
'''   ' Restaura el control.
'''   For I = myFlex.FixedRows To myFlex.Rows - 1
'''      If I < myFlex.Row Or I > myFlex.RowSel Then
'''         myFlex.RowHeight(I) = myFlex.RowData(I)
'''      End If
'''   Next
'''   For I = myFlex.FixedCols To myFlex.Cols - 1
'''      If I < myFlex.Col Or I > myFlex.ColSel Then
'''         myFlex.ColWidth(I) = myFlex.ColData(I)
'''      End If
'''   Next
'''   myFlex.TopRow = tr
'''   myFlex.LeftCol = lc
'''   myFlex.HighLight = hl
'''   myFlex.Redraw = True
'''End Sub

Private Sub Command1_Click()
  RatonReloj
  DGQuery.Visible = False
  ImprimirCatalogoInv AdoQuery
''  'SelectMSFGrid MSHGQuery, AdoQuery, sSQL
''  Select_Adodc_Grid DGQuery, AdoQuery, sSQL
  DGQuery.Visible = True
  RatonNormal
End Sub

Private Sub Command2_Click()
  Unload CatalogoCtas
End Sub

Private Sub Command3_Click()
  Codigo = CambioCodigoKardex(MBoxCtaI.Text)
  Codigo1 = CambioCodigoKardex(MBoxCtaF.Text)
  ListarCatalogoInventario
  SQLMsg1 = "PRODUCTOS DE INVENTARIO"
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto CatalogoCtas, AdoQuery
     DGQuery.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyF5 Then
     DGQuery.AllowUpdate = True
     MsgBox "Proceso Aceptado, puede Modificar"
     DGQuery.SetFocus
  End If
End Sub

Private Sub Form_Activate()
  RatonReloj
 'Maximizamos segun la resolucion de la pantalla
  DGQuery.Height = MDI_Y_Max - DGQuery.Top - 400
  DGQuery.width = MDI_X_Max - 100
  AdoQuery.Top = DGQuery.Top + DGQuery.Height
  
  Codigo = Ninguno: Codigo1 = Ninguno
  FormatoCodigoKardex MBoxCtaI
  FormatoCodigoKardex MBoxCtaF
  ListarCatalogoInventario
  SQLMsg1 = "PRODUCTOS DE INVENTARIO"
  RatonNormal
End Sub

Private Sub Form_Load()
  'CentrarForm CatalogoCtas
  ConectarAdodc AdoQuery
End Sub

Public Sub ListarCatalogoInventario()
  Codigo = CambioCodigoKardex(MBoxCtaI.Text)
  Codigo1 = CambioCodigoKardex(MBoxCtaF.Text)
  If Codigo = "" Then Codigo = Ninguno
  If Codigo1 = "" Then Codigo1 = Ninguno
  sSQL = "SELECT TC,Codigo_Inv,Producto,PVP,Codigo_Barra,Cta_Inventario,Unidad,Cantidad," _
       & "Cta_Costo_Venta,Cta_Ventas,Cta_Ventas_0,Cta_Ventas_Ant,Cta_Venta_Anticipada," _
       & "IVA,INV,Codigo_IESS, Codigo_RES, Marca,Reg_Sanitario,Ayuda " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If Codigo <> Ninguno And Codigo1 <> Ninguno Then
     sSQL = sSQL & "AND Codigo_Inv BETWEEN '" & Codigo & "' and '" & Codigo1 & "' "
  End If
  If CheqPM.value = 1 Then sSQL = sSQL & "AND TC = 'P' "
  sSQL = sSQL & "ORDER BY Codigo_Inv "
  Select_Adodc_Grid DGQuery, AdoQuery, sSQL
End Sub

