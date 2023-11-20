VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{64AED23E-31A2-4023-8C7D-E628B15843D8}#1.0#0"; "Code39X.ocx"
Begin VB.Form CatalogoActivos 
   Caption         =   "PRESENTACION DEL CATALOGO DE CUENTA"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmNewResp 
      Caption         =   "SELECCIONE A QUI EL NUEVO RESPONSABLE:"
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
      Left            =   105
      TabIndex        =   15
      Top             =   6405
      Visible         =   0   'False
      Width           =   4740
      Begin MSDataListLib.DataCombo DCNewResp 
         Bindings        =   "Catalogo.frx":0000
         DataSource      =   "AdoNewResp"
         Height          =   315
         Left            =   105
         TabIndex        =   16
         Top             =   315
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
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
   End
   Begin VB.OptionButton OpcFiltro 
      Caption         =   "Presentar por filtro de busqueda"
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
      Left            =   5040
      TabIndex        =   14
      Top             =   105
      Width           =   3270
   End
   Begin VB.OptionButton OpcRango 
      Caption         =   "Presentar por Rango de Códigos"
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
      Left            =   1155
      TabIndex        =   13
      Top             =   105
      Value           =   -1  'True
      Width           =   3165
   End
   Begin MSDataListLib.DataCombo DCFiltros 
      Bindings        =   "Catalogo.frx":0019
      DataSource      =   "AdoFiltros"
      Height          =   315
      Left            =   5355
      TabIndex        =   12
      Top             =   1260
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
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
   Begin VB.ListBox LstCampos 
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
      Left            =   5355
      TabIndex        =   11
      Top             =   420
      Width           =   4425
   End
   Begin MSDataGridLib.DataGrid DGQuery 
      Bindings        =   "Catalogo.frx":0032
      Height          =   4860
      Left            =   1155
      TabIndex        =   7
      ToolTipText     =   "<F1> Genera Archivo,  <Ctrl>+<R> Cambiar el Responsable, <Ctrl>+<Del> "
      Top             =   1680
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   8573
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1365
      ScaleHeight     =   855
      ScaleWidth      =   3480
      TabIndex        =   10
      Top             =   3000
      Width           =   3480
   End
   Begin VB.CommandButton Command4 
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
      Left            =   105
      Picture         =   "Catalogo.frx":0049
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1785
      Width           =   960
   End
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
      Left            =   1470
      TabIndex        =   8
      Top             =   1260
      Width           =   3060
   End
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   1155
      Top             =   6825
      Width           =   4320
      _ExtentX        =   7620
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
      Left            =   105
      Picture         =   "Catalogo.frx":048B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   960
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
      Left            =   105
      Picture         =   "Catalogo.frx":08CD
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2625
      Width           =   960
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
      Left            =   105
      Picture         =   "Catalogo.frx":1197
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   945
      Width           =   960
   End
   Begin MSMask.MaskEdBox MBoxCtaF 
      Height          =   330
      Left            =   3150
      TabIndex        =   3
      Top             =   840
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
      Left            =   3150
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
   Begin Code39X.Code39Clt Code39Clt1 
      Left            =   105
      Top             =   3465
      _ExtentX        =   1905
      _ExtentY        =   1085
   End
   Begin MSAdodcLib.Adodc AdoFiltros 
      Height          =   330
      Left            =   1365
      Top             =   5460
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
      Caption         =   "Filtros"
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
   Begin MSAdodcLib.Adodc AdoPatron2 
      Height          =   330
      Left            =   1365
      Top             =   5880
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
      Caption         =   "Patron2"
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
   Begin MSAdodcLib.Adodc AdoNewResp 
      Height          =   330
      Left            =   1365
      Top             =   5145
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
      Caption         =   "NewResp"
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
      Left            =   1470
      TabIndex        =   0
      Top             =   420
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
      Left            =   1470
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "CatalogoActivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ColocarCodigoBarra(CodigoDeBarra As String)
Dim PosSup, PosIzq As Single
  Code39Clt1.AlturaBarra = 45
  Code39Clt1.TamBarra = 1
  Code39Clt1.ColorCodigo = "N"
  Code39Clt1.ValorCodigo = CodigoDeBarra
  Code39Clt1.RealizarCodigo
  
  Picture1 = Clipboard.GetData
  Picture1.FontBold = True
  PosIzq = ((Picture1.width - Picture1.TextWidth(Code39Clt1.ValorCodigo)) / 2)
  PosSup = Picture1.Height - 360
  If PosIzq < 0 Then PosIzq = 0.1
  If PosSup < 0 Then PosSup = 0.1
  Picture1.Line (150, PosSup + 140)-(Picture1.width - 150, Picture1.Height), QBColor(Blanco_Brillante), BF
  Picture1.CurrentX = PosIzq
  Picture1.CurrentY = PosSup + 150
  Picture1.Print Code39Clt1.ValorCodigo
  PosSup = Picture1.Height - 160
''  Cadena = "USD$ " & Format(PVP, "#,##0.0000")
''  PosIzq = ((Picture1.Width - Picture1.TextWidth(Cadena)) / 2)
''  Picture1.CurrentX = PosIzq
''  Picture1.CurrentY = PosSup
''  Picture1.Print Cadena
  Picture1.FontBold = False
  If TxtBarra.Text = Ninguno Then TxtBarra.Text = Code39Clt1.ValorCodigo
End Sub

Private Sub CopySelectedPictureToClipboard(myFlex As MSHFlexGrid)
   Dim I As Integer, tr As Long, lc As Long, hl As Integer
   ' Se prepara para funcionar.
   myFlex.Redraw = False  ' Para eliminar el parpadeo.
   hl = myFlex.HighLight  ' Guarda los valores actuales.
   tr = myFlex.TopRow
   lc = myFlex.LeftCol
   myFlex.HighLight = 0  ' No se resalta la imagen.
   ' Oculta las filas y columnas no seleccionadas.
   ' (Guarda los tamaños originales en las propiedades
   ' RowData y ColData).
   For I = myFlex.FixedRows To myFlex.Rows - 1
      If I < myFlex.Row Or I > myFlex.RowSel Then
         myFlex.RowData(I) = myFlex.RowHeight(I)
         myFlex.RowHeight(I) = 0
      End If
   Next
   For I = myFlex.FixedCols To myFlex.Cols - 1
      If I < myFlex.Col Or I > myFlex.ColSel Then
         myFlex.ColData(I) = myFlex.ColWidth(I)
         myFlex.ColWidth(I) = 0
      End If
   Next
   ' Se desplaza a la esquina superior izquierda.
   myFlex.TopRow = myFlex.FixedRows
   myFlex.LeftCol = myFlex.FixedCols
   ' Copia la imagen.
   Clipboard.Clear
   On Error Resume Next
   myFlex.PictureType = 0 ' Color.
   Clipboard.SetData myFlex.Picture
   If Error <> 0 Then
      myFlex.PictureType = 1 ' Monocromo.
      Clipboard.SetData myFlex.Picture
   End If
   ' Restaura el control.
   For I = myFlex.FixedRows To myFlex.Rows - 1
      If I < myFlex.Row Or I > myFlex.RowSel Then
         myFlex.RowHeight(I) = myFlex.RowData(I)
      End If
   Next
   For I = myFlex.FixedCols To myFlex.Cols - 1
      If I < myFlex.Col Or I > myFlex.ColSel Then
         myFlex.ColWidth(I) = myFlex.ColData(I)
      End If
   Next
   myFlex.TopRow = tr
   myFlex.LeftCol = lc
   myFlex.HighLight = hl
   myFlex.Redraw = True
End Sub

Private Sub Command1_Click()
  RatonReloj
  DGQuery.Visible = False
  Codigo = CambioCodigoKardex(MBoxCtaI)
  Codigo1 = CambioCodigoKardex(MBoxCtaF)
  ListarCatalogoInventario True
  SQLMsg1 = "ACTIVOS FIJOS"
  ImprimirCatalogoActivos AdoQuery
  Codigo = CambioCodigoKardex(MBoxCtaI)
  Codigo1 = CambioCodigoKardex(MBoxCtaF)
  ListarCatalogoInventario
  DGQuery.Visible = True
  RatonNormal
End Sub

Private Sub Command2_Click()
  Unload CatalogoActivos
End Sub

Private Sub Command3_Click()
  Codigo = CambioCodigoKardex(MBoxCtaI)
  Codigo1 = CambioCodigoKardex(MBoxCtaF)
  ListarCatalogoInventario
  SQLMsg1 = "ACTIVOS FIJOS"
End Sub

Private Sub Command4_Click()
  IE = Val(InputBox("Cantidad de Etiquetas por Fila", "IMPRESION CODIGO DE BARRAS DE ACTIVOS", "4"))
  If IE > 0 Then
     DGQuery.Visible = False
     Imprimir_Codigos_De_Activos AdoQuery, Code39Clt1, IE, Picture1
     DGQuery.Visible = True
  End If
End Sub

Private Sub DCNewResp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCNewResp_LostFocus()
  Codigo1 = Ninguno
  Codigo2 = Ninguno
  With AdoQuery.Recordset
   If .RecordCount > 0 Then Codigo1 = .Fields("Responsable")
  End With
  With AdoNewResp.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCNewResp.Text & "' ")
       If Not .EOF Then Codigo2 = .Fields("CI_RUC")
   End If
  End With
  If Codigo2 <> Ninguno Then
     sSQL = "UPDATE Catalogo_Productos " _
          & "SET Responsable = '" & Codigo2 & "' " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Responsable = '" & Codigo1 & "' " _
          & "AND TC = 'P' "
     If OpcRango.value Then
        If Codigo <> Ninguno And Codigo1 <> Ninguno Then
           sSQL = sSQL & "AND Codigo_Activo BETWEEN '" & Codigo & "' and '" & Codigo1 & "' "
        End If
     Else
        With AdoQuery.Recordset
            Select Case .Fields(LstCampos.ListIndex).Type
              Case TadBoolean
                   sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = " & CBool(DCFiltros.Text) & " "
              Case TadDate, TadDate1
                   sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = #" & CDate(DCFiltros.Text) & "# "
              Case TadTime
                   sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = " & DCFiltros.Text & " "
              Case TadByte, TadInteger, TadLong
                   sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = " & Val(DCFiltros.Text) & " "
              Case TadSingle, TadDouble, TadCurrency
                   sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = " & Val(DCFiltros.Text) & " "
              Case TadText
                   sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = '" & DCFiltros.Text & "' "
              Case Else
                   sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = '" & DCFiltros.Text & "' "
            End Select
        End With
     End If
     ConectarAdoExecute sSQL
     MsgBox "Cambio Realizado con exito, Vuelva a Consultar"
     Unload CatalogoActivos
  End If
End Sub

Private Sub DGQuery_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  If KeyCode = vbKeyF1 Then
     DGQuery.Visible = False
     GenerarDataTexto CatalogoActivos, AdoQuery
     DGQuery.Visible = True
  End If
  If CtrlDown And KeyCode = vbKeyF5 Then
     DGQuery.AllowUpdate = True
     MsgBox "Proceso Aceptado, puede Modificar"
     DGQuery.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyR Then
     FrmNewResp.Left = 1600
     FrmNewResp.Top = 1600
     FrmNewResp.Visible = True
     DCNewResp.SetFocus
  End If
  If CtrlDown And KeyCode = vbKeyDelete Then
     Codigo = Ninguno
     With AdoQuery.Recordset
      If .RecordCount > 0 Then Codigo = .Fields("Codigo_Activo")
     End With
     sSQL = "DELETE * " _
          & "FROM Catalogo_Productos " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND Codigo_Activo = '" & Codigo & "' "
     ConectarAdoExecute sSQL
     MsgBox "Proceso terminado con exito, Vuelva a consultar"
  End If
End Sub

Private Sub Form_Activate()
  RatonReloj
  Codigo = Ninguno: Codigo1 = Ninguno
  FormatoCodigoKardex MBoxCtaI
  FormatoCodigoKardex MBoxCtaF
  ListarCatalogoInventario
  LstCampos.Clear
  With AdoQuery.Recordset
   If .RecordCount > 0 Then
       For I = 0 To .Fields.Count - 1
           LstCampos.AddItem .Fields(I).Name
       Next I
   End If
  End With
  sSQL = "SELECT CI_RUC,Cliente " _
       & "FROM Clientes " _
       & "WHERE TD <> 'R' " _
       & "ORDER BY Cliente "
  SelectDBCombo DCNewResp, AdoNewResp, sSQL, "Cliente"
  SQLMsg1 = "ACTIVOS FIJOS"
  RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoQuery
  ConectarAdodc AdoFiltros
  ConectarAdodc AdoNewResp
  
  DGQuery.Height = MDI_Y_Max - DGQuery.Top - 300
  DGQuery.width = MDI_X_Max - DGQuery.Left
  AdoQuery.Top = DGQuery.Top + DGQuery.Height + 50
  
  ''Label1.Top = DGQuery.Top + DGQuery.Height + 50
End Sub

Public Sub ListarCatalogoInventario(Optional Solo_Imp As Boolean)
  Codigo = CambioCodigoKardex(MBoxCtaI)
  Codigo1 = CambioCodigoKardex(MBoxCtaF)
  If Codigo = "" Then Codigo = Ninguno
  If Codigo1 = "" Then Codigo1 = Ninguno
  sSQL = "UPDATE Catalogo_Productos " _
       & "SET Nombre_Responsable = Responsable " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' "
  If Codigo <> Ninguno And Codigo1 <> Ninguno Then
     sSQL = sSQL & "AND Codigo_Activo BETWEEN '" & Codigo & "' and '" & Codigo1 & "' "
  End If
  ConectarAdoExecute sSQL
  
  If SQL_Server Then
     sSQL = "UPDATE Catalogo_Productos " _
          & "SET Nombre_Responsable = C.Cliente " _
          & "FROM Catalogo_Productos As CA,Clientes As C "
  Else
     sSQL = "UPDATE Catalogo_Productos As CA,Clientes As C " _
          & "SET CA.Nombre_Responsable = C.Cliente "
  End If
  sSQL = sSQL & "WHERE CA.Item = '" & NumEmpresa & "' " _
       & "AND CA.Periodo = '" & Periodo_Contable & "' " _
       & "AND CA.TC = 'P' " _
       & "AND CA.Responsable = C.CI_RUC "
  If Codigo <> Ninguno And Codigo1 <> Ninguno Then
     sSQL = sSQL & "AND CA.Codigo_Activo BETWEEN '" & Codigo & "' and '" & Codigo1 & "' "
  End If
  ConectarAdoExecute sSQL

  If Solo_Imp Then
     sSQL = "SELECT Producto,Codigo_Barra,Total_Compra As Valor,Tipo,Ubicacion,Nombre_Responsable "
  Else
     sSQL = "SELECT Codigo_Activo,Producto,Fecha_Compra,Factura_No,Cantidad,Sub_Total,Total_IVA,Total_Compra,Tipo,Ubicacion,Detalle,Codigo_Barra,Nombre_Responsable,Responsable "
  End If
  sSQL = sSQL & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  If OpcRango.value Then
     If Codigo <> Ninguno And Codigo1 <> Ninguno Then
        sSQL = sSQL & "AND Codigo_Activo BETWEEN '" & Codigo & "' and '" & Codigo1 & "' "
     End If
  Else
     With AdoQuery.Recordset
          Select Case .Fields(LstCampos.ListIndex).Type
            Case TadBoolean
                 sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = " & CBool(DCFiltros.Text) & " "
            Case TadDate, TadDate1
                 sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = #" & CDate(DCFiltros.Text) & "# "
            Case TadTime
                 sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = " & DCFiltros.Text & " "
            Case TadByte, TadInteger, TadLong
                 sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = " & Val(DCFiltros.Text) & " "
            Case TadSingle, TadDouble, TadCurrency
                 sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = " & Val(DCFiltros.Text) & " "
            Case TadText
                 sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = '" & DCFiltros.Text & "' "
            Case Else
                 sSQL = sSQL & "AND " & .Fields(LstCampos.ListIndex).Name & " = '" & DCFiltros.Text & "' "
          End Select
     End With
  End If
  If CheqPM.value = 1 Then sSQL = sSQL & "AND TC = 'P' "
  sSQL = sSQL & "ORDER BY Ubicacion,Producto,Codigo_Activo "
  SelectDataGrid DGQuery, AdoQuery, sSQL
End Sub

Private Sub LstCampos_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub LstCampos_LostFocus()
  sSQL = "SELECT " & LstCampos.Text & " " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC <> 'I' " _
       & "GROUP BY " & LstCampos.Text & " " _
       & "ORDER BY " & LstCampos.Text & " "
  SelectDBCombo DCFiltros, AdoFiltros, sSQL, LstCampos.Text
  DCFiltros.SetFocus
End Sub
