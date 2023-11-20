VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form ListaStock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PRODUCTOS"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DGStock 
      Bindings        =   "FListaStock.frx":0000
      Height          =   4110
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   7250
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
   Begin MSAdodcLib.Adodc AdoStock 
      Height          =   330
      Left            =   105
      Top             =   4305
      Width           =   9570
      _ExtentX        =   16880
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
      Caption         =   "AdoStock"
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   10815
      TabIndex        =   2
      Top             =   4305
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   9765
      TabIndex        =   1
      Top             =   4305
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoMarca 
      Height          =   330
      Left            =   210
      Top             =   2835
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "Marca"
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
Attribute VB_Name = "ListaStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Encontro As Boolean
  If AdoStock.Recordset.RecordCount > 0 Then
     Encontro = Leer_Codigo_Inv(DGStock.Columns(2).Text, FechaSistema, Cod_Bodega)
    ' MsgBox "=> " & DatInv.Codigo_Inv
  Else
     DatInv.Codigo_Inv = Ninguno
  End If
  Unload Me
End Sub

Private Sub Command2_Click()
  DatInv.Codigo_Inv = Ninguno
  Unload Me
End Sub

Private Sub DGStock_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Activate()
Dim Medicar As String
    
    sSQL = "DELETE * " _
         & "FROM Catalogo_Marcas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    ConectarAdoExecute sSQL
    
    sSQL = "INSERT INTO Catalogo_Marcas (CodMar, Marca, Item, Periodo) " _
         & "SELECT Codigo_Inv, Mid$(Producto,1,30), Item, Periodo " _
         & "FROM Catalogo_Productos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'I' "
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Catalogo_Productos " _
         & "SET X = '.' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Catalogo_Productos " _
         & "SET X = 'B' " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'P' " _
         & "AND INV <> " & Val(adFalse) & " " _
         & "AND LEN(Cta_Inventario) > 2 "
     If UCase$(Mid$(DatInv.Patron_Busqueda, 1, 2)) = "M:" Then
        Medicar = UCase$(Mid$(DatInv.Patron_Busqueda, 3, Trim$(Len(DatInv.Patron_Busqueda))))
        sSQL = sSQL & "AND Detalle LIKE '%" & Medicar & "%' "
     Else
        sSQL = sSQL & "AND Producto LIKE '%" & DatInv.Patron_Busqueda & "%' "
     End If
    ConectarAdoExecute sSQL
    
    sSQL = "UPDATE Catalogo_Productos " _
         & "SET Stock_Actual = (SELECT SUM(Entrada)-SUM(Salida) " _
         & "                    FROM Trans_Kardex As TK " _
         & "                    WHERE TK.Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
         & "                    AND TK.T <> 'A' " _
         & "                    AND TK.Item = Catalogo_Productos.Item " _
         & "                    AND TK.Periodo = Catalogo_Productos.Periodo " _
         & "                    AND TK.Codigo_Inv = Catalogo_Productos.Codigo_Inv) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'P' " _
         & "AND X = 'B' "
    ConectarAdoExecute sSQL
        
    sSQL = "UPDATE Catalogo_Productos " _
         & "SET Salidas = (SELECT SUM(Cantidad) " _
         & "                      FROM Detalle_Factura As DF " _
         & "                      WHERE DF.Fecha <= #" & BuscarFecha(FechaSistema) & "# " _
         & "                      AND DF.T <> 'A' " _
         & "                      AND DF.Cantidad > 0 " _
         & "                      AND DF.C = " & Val(adFalse) & " " _
         & "                      AND DF.Item = Catalogo_Productos.Item " _
         & "                      AND DF.Item = Catalogo_Productos.Item " _
         & "                      AND DF.Periodo = Catalogo_Productos.Periodo " _
         & "                      AND DF.Codigo = Catalogo_Productos.Codigo_Inv) " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'P' " _
         & "AND X = 'B' "
    ConectarAdoExecute sSQL
    
   'Enceramos todos los nulos
    sSQL = "UPDATE Catalogo_Productos " _
         & "SET Stock_Actual = 0 " _
         & "WHERE Stock_Actual IS NULL " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'P' " _
         & "AND X = 'B' "
    ConectarAdoExecute sSQL
        
    sSQL = "UPDATE Catalogo_Productos " _
         & "SET Salidas = 0 " _
         & "WHERE Salidas IS NULL " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'P' " _
         & "AND X = 'B' "
    ConectarAdoExecute sSQL
              
    sSQL = "UPDATE Catalogo_Productos " _
         & "SET Stock_Actual = Stock_Actual - Salidas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'P' " _
         & "AND X = 'B' "
    ConectarAdoExecute sSQL
    
    If SQL_Server Then
       sSQL = "UPDATE Catalogo_Productos " _
            & "SET Marca = CM.Marca " _
            & "FROM Catalogo_Productos As CP, Catalogo_Marcas As CM "
    Else
       sSQL = "UPDATE Catalogo_Productos As CP, Catalogo_Marcas As CM " _
            & "SET CP.Marca = CM.Marca "
    End If
    sSQL = sSQL _
         & "WHERE CP.Item = '" & NumEmpresa & "' " _
         & "AND CP.Periodo = '" & Periodo_Contable & "' " _
         & "AND CP.TC = 'P' " _
         & "AND CP.Item = CM.Item " _
         & "AND CP.Periodo = CM.Periodo " _
         & "AND CP.Codigo_Sup = CM.CodMar "
    ConectarAdoExecute sSQL
        
    sSQL = "SELECT Marca,TC,Codigo_Inv,Producto,PVP,Stock_Actual,Mid$(Ubicacion,1,10) As Lugar " _
         & "FROM Catalogo_Productos " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND Stock_Actual > 0 " _
         & "AND X = 'B' " _
         & "ORDER BY Codigo_Inv "
    SelectDataGrid DGStock, AdoStock, sSQL
End Sub

Private Sub Form_Load()
    ConectarAdodc AdoStock
    ConectarAdodc AdoMarca
End Sub
