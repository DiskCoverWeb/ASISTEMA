VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FStockInventario 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SUBIDA DE INVENTARIO"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "StockInv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Recibir &Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1470
      Picture         =   "StockInv.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   105
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFFF&
      Height          =   3795
      Left            =   3255
      TabIndex        =   8
      Top             =   1050
      Width           =   2535
   End
   Begin ComctlLib.ProgressBar ProgBarra 
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   4935
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
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
      Height          =   315
      Left            =   105
      TabIndex        =   4
      Top             =   1050
      Width           =   3060
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
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
      Height          =   645
      Left            =   3255
      Picture         =   "StockInv.frx":0CE8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
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
      Height          =   3465
      Left            =   105
      TabIndex        =   5
      Top             =   1365
      Width           =   3060
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   525
      Top             =   1890
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
   Begin MSAdodcLib.Adodc AdoQuery 
      Height          =   330
      Left            =   525
      Top             =   2205
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
   Begin MSAdodcLib.Adodc AdoAct 
      Height          =   330
      Left            =   525
      Top             =   2520
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Act"
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
   Begin MSMask.MaskEdBox MBFechaI 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   65535
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
   Begin MSAdodcLib.Adodc AdoInv 
      Height          =   330
      Left            =   525
      Top             =   2835
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Act"
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
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARCHIVO:"
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
      Left            =   3255
      TabIndex        =   6
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &FECHA ING."
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
      Width           =   1275
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &ORIGEN"
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
      Width           =   3060
   End
End
Attribute VB_Name = "FStockInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AdoStrCnnOld As String
Dim AdoStrCnn1 As String
Dim NumFile As Integer
Dim RutaGeneraFile As String
Dim XAdoStrCnn As String
Dim IJ As Long
Dim ModuloResp As String
Dim RetVal

Public Sub TipoProcesos(Opciones As String)
  NombreArchivo = ""
  Select Case Opciones
    Case "DESCUENT"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case "FACTURAS"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case "NOMINA"
         Dir1.Path = RutaBackup & "\" & Opciones & "\"
         File1.FileName = Dir1.Path & "\*.*"
    Case Else
         Dir1.Path = RutaBackup & "\"
         File1.FileName = Dir1.Path & "\*.*"
  End Select
  Dir1.Refresh
End Sub

Private Sub Command1_Click()
Dim AuxNumEmp As String
Dim DiaV As Integer
Dim MesV As Integer
Dim AñoV As Integer
Dim Total_Alumnos As Long
Dim CamposFile() As Campos_Tabla
Dim EsComa As Boolean
Dim EsTab As Boolean
  TextoImprimio = ""
  EsComa = False
  EsTab = True
  FechaValida MBFechaI
  Mifecha = BuscarFecha(MBFechaI)
  FechaTexto = MBFechaI ' FechaSistema
  DiarioCaja = 1
  RutaGeneraFile = UCase(Dir1.Path & "\" & NombreArchivo)
  TotalIngreso = 0
  Real1 = 0
  Real2 = 0
  Contador = 0: FileResp = 0
  ProgBarra.Value = 0
  ProgBarra.Min = 0
 'Establecemos los campos del archivo plano del Banco
  NumFile = FreeFile
  Total_Alumnos = 0
  Open RutaGeneraFile For Input As #NumFile
       Do While Not EOF(NumFile)
          Line Input #NumFile, Cod_Field
          If Total_Alumnos = 1 Then
             Cadena = Cod_Field
             FechaTexto = ""
             TotalReg = Len(Cadena)
             No_Desde = 1: No_Hasta = 1
             CantCampos = 0
             Do While Len(Cadena) > 0
                Do
                  No_Hasta = No_Hasta + 1
                  If Mid(Cadena, No_Hasta, 1) = vbTab Then EsTab = True
                Loop Until Mid(Cadena, No_Hasta, 1) = vbTab Or No_Hasta > TotalReg
               'Obtenemos la fecha de subida
                If CantCampos = 8 Then FechaTexto = Trim(Mid(Cadena, No_Desde + 1, No_Hasta - 2))
                CantCampos = CantCampos + 1
                Cadena = Mid(Cadena, No_Hasta, Len(Cadena))
                TotalReg = Len(Cadena)
                No_Desde = 1: No_Hasta = 1
             Loop
             ReDim CamposFile(CantCampos + 1) As Campos_Tabla
             For I = 0 To CantCampos
                 CamposFile(I).Campo = "C" & Format(I, "00")
             Next I
          End If
          Total_Alumnos = Total_Alumnos + 1
       Loop
  Close #NumFile
 'Eliminamos los abonos de este dia
  Mifecha = BuscarFecha(FechaTexto)
  Numero = Mid(FechaTexto, 7, 4) & Mid(FechaTexto, 4, 2) & Mid(FechaTexto, 1, 2)
  sSQL = "DELETE * " _
       & "FROM Trans_Kardex " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP = 'CD' " _
       & "AND Numero = " & Numero & " "
  ConectarAdoExecute sSQL
  
  sSQL = "DELETE * " _
       & "FROM Comprobantes " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP = 'CD' " _
       & "AND Numero = " & Numero & " "
  ConectarAdoExecute sSQL
  
  sSQL = "DELETE * " _
       & "FROM Transacciones " _
       & "WHERE Fecha = #" & Mifecha & "# " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TP = 'CD' " _
       & "AND Numero = " & Numero & " "
  ConectarAdoExecute sSQL
  'MsgBox FechaTexto & vbCrLf & sSQL
  ProgBarra.Max = Total_Alumnos + 1
  
 'Consultamos los Codigos a verificar
  sSQL = "SELECT * " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "ORDER BY Codigo_Inv "
  SelectAdodc AdoInv, sSQL
' Comenzamos a leer el archivo de Abonos
  With AdoInv.Recordset
   If .RecordCount > 0 Then
       RutaGeneraFile = UCase(Dir1.Path & "\" & NombreArchivo)
       Contador = 0: FileResp = 0
       NumFile = FreeFile
       Open RutaGeneraFile For Input As #NumFile
         Line Input #NumFile, Cod_Field
         'MsgBox Cod_Field
         ProgBarra.Value = ProgBarra.Value + 1
         Do While Not EOF(NumFile)
            Line Input #NumFile, Cod_Field
            ProgBarra.Value = ProgBarra.Value + 1
            No_Desde = 1: No_Hasta = 1
            Cadena = Cod_Field
            I = 0
            TotalReg = Len(Cadena)
            Do While Len(Cadena) > 0
               Do
                 No_Hasta = No_Hasta + 1
               Loop Until Mid(Cadena, No_Hasta, 1) = "," Or Mid(Cadena, No_Hasta, 1) = vbTab Or No_Hasta > TotalReg
               'MsgBox CamposFile(I).Campo
               CamposFile(I).Valor = Mid(Cadena, No_Desde, No_Hasta)
               If Len(CamposFile(I).Valor) > 1 Then
                  If EsTab Then
                     If Mid(CamposFile(I).Valor, 1, 1) = vbTab Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 2, Len(CamposFile(I).Valor))
                     If Mid(CamposFile(I).Valor, Len(CamposFile(I).Valor), 1) = vbTab Then CamposFile(I).Valor = Mid(CamposFile(I).Valor, 1, Len(CamposFile(I).Valor) - 1)
                  End If
               End If
               I = I + 1
               Cadena = Mid(Cadena, No_Hasta, Len(Cadena))
               TotalReg = Len(Cadena)
               No_Desde = 1: No_Hasta = 1
            Loop
          ' Obtenemos el Valor de Abonos de Pensiones o Matriculas
            Codigo = CamposFile(1).Valor & CamposFile(0).Valor
            CodigoP = CamposFile(0).Valor
            Producto = CamposFile(2).Valor
            Entrada = Val(CamposFile(3).Valor)
            Costo = Val(CamposFile(4).Valor)
            Precio = Val(CamposFile(6).Valor)
            Total_IVA = Val(CamposFile(10).Valor)
            Cta = CamposFile(12).Valor
            Cta1 = CamposFile(13).Valor
            
            CodigoInv = Codigo
            Total = Entrada * Costo
           .MoveFirst
           .Find ("Codigo_Inv = '" & Codigo & "' ")
            If Not .EOF Then
               CodigoInv = .Fields("Codigo_Inv")
            Else
               CodigoInv = Codigo
              .AddNew
               SetFields AdoInv, "Codigo_Inv", CodigoInv
            End If
            'MsgBox Producto
            SetFields AdoInv, "Producto", Mid(Producto, 1, 40)
            If Total_IVA > 0 Then SetFields AdoInv, "IVA", adTrue
            SetFields AdoInv, "PVP", Precio
            SetFields AdoInv, "Codigo_Barra", CodigoP
            SetFields AdoInv, "TC", "P"
            SetFields AdoInv, "Cta_Inventario", CamposFile(12).Valor
            SetFields AdoInv, "Cta_Costo_Venta", CamposFile(14).Valor
            SetFields AdoInv, "Cta_Ventas", CamposFile(15).Valor
            SetUpdate AdoInv
            'MsgBox FechaTexto & vbCrLf & Numero
            If Total > 0 Then
               SetAdoAddNew "Trans_Kardex"
               SetAdoFields "T", Normal
               SetAdoFields "TP", "CD"
               SetAdoFields "Fecha", FechaTexto
               SetAdoFields "Numero", Numero
               SetAdoFields "CodigoC", Ninguno
               SetAdoFields "Codigo_Inv", CodigoInv
               SetAdoFields "Codigo_Barra", CodigoP
               SetAdoFields "Orden_No", Numero
               SetAdoFields "Entrada", Entrada
               SetAdoFields "Valor_Unitario", Costo
               SetAdoFields "Valor_Total", Total
               SetAdoFields "Costo", Costo
               SetAdoFields "PVP", Precio
               SetAdoFields "Total", Total
               SetAdoFields "Total_IVA", Total_IVA
               SetAdoFields "Codigo_P", Ninguno
               SetAdoFields "Cta_Inv", Cta
               SetAdoFields "Contra_Cta", Cta1
               SetAdoFields "Kardex", Contador + 1
               SetAdoUpdate
            End If
            If CamposFile(9).Valor = "1" Then
               Real1 = Real1 + Total
               Cta1 = CamposFile(13).Valor
            Else
               Real2 = Real2 + Total
               Cta_Aux = CamposFile(13).Valor
            End If
            TotalIngreso = TotalIngreso + Total
            Contador = Contador + 1
            If Contador > ProgBarra.Max Then Contador = ProgBarra.Max
            ProgBarra.Value = Contador
         Loop
         ProgBarra.Value = ProgBarra.Max
       Close #NumFile
       If TotalIngreso > 0 Then
          SetAdoAddNew "Transacciones"
          SetAdoFields "T", Normal
          SetAdoFields "TP", "CD"
          SetAdoFields "Fecha", FechaTexto
          SetAdoFields "Numero", Numero
          SetAdoFields "Cta", Cta
          SetAdoFields "Debe", TotalIngreso
          SetAdoFields "Haber", 0
          SetAdoFields "ID", Contador + 1
          SetAdoUpdate
          If Real1 > 0 Then
             SetAdoAddNew "Transacciones"
             SetAdoFields "T", Normal
             SetAdoFields "TP", "CD"
             SetAdoFields "Fecha", FechaTexto
             SetAdoFields "Numero", Numero
             SetAdoFields "Cta", Cta1
             SetAdoFields "Debe", 0
             SetAdoFields "Haber", Real1
             SetAdoFields "ID", Contador + 2
             SetAdoUpdate
          End If
          If Real2 > 0 Then
             SetAdoAddNew "Transacciones"
             SetAdoFields "T", Normal
             SetAdoFields "TP", "CD"
             SetAdoFields "Fecha", FechaTexto
             SetAdoFields "Numero", Numero
             SetAdoFields "Cta", Cta_Aux
             SetAdoFields "Debe", 0
             SetAdoFields "Haber", Real2
             SetAdoFields "ID", Contador + 3
             SetAdoUpdate
          End If
          SetAdoAddNew "Comprobantes"
          SetAdoFields "T", Normal
          SetAdoFields "TP", "CD"
          SetAdoFields "Fecha", FechaTexto
          SetAdoFields "Numero", Numero
          SetAdoFields "Codigo_B", Ninguno
          SetAdoFields "Concepto", "Ingreso de Mercaderia por Archivo del día " & FechaTexto
          SetAdoFields "Monto_Total", TotalIngreso
          SetAdoUpdate
       End If
      'Consultamos los Codigos a verificar
'''       sSQL = "UPDATE Catalogo_Productos " _
'''            & "SET INV = " & Val(adTrue) & " " _
'''            & "WHERE Item = '" & NumEmpresa & "' " _
'''            & "AND Periodo = '" & Periodo_Contable & "' " _
'''            & "AND INV = " & Val(adFalse) & " " _
'''            & "AND TC = 'P' "
'''       ConectarAdoExecute sSQL
   Else
       MsgBox "No Existen Facturas Pendientes"
   End If
  End With
  RatonNormal
  ProgBarra.Value = ProgBarra.Max
  FStockInventario.Caption = "INGRESO DE INVNETARIO"
  MsgBox "Proceso Terminado del Archivo por " & TotalIngreso & ": " & vbCrLf & UCase(Dir1.Path & "\" & NombreArchivo)
'  If TextoImprimio <> "" Then FInfoError.Show
End Sub

Private Sub Command2_Click()
  Unload FStockInventario
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
  SiguienteControl
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
  NombreArchivo = File1.FileName
  If KeyCode = vbKeyDelete Then
     Mensajes = "Esta seguro de Eliminar: " & File1.FileName
     Titulo = "Pregunta de Eliminacion"
     If BoxMensaje = vbYes Then Kill File1.Path & "\" & File1.FileName
     File1.FileName = Dir1.Path & "\*.*"
  End If
End Sub

Private Sub File1_LostFocus()
  NombreArchivo = UCase(File1.FileName)
End Sub

Private Sub Form_Activate()
  FechaValida MBFechaI
  
  Drive1.Drive = Mid(RutaSysBases, 1, 2)
  RatonNormal
  RutaBackup = RutaSysBases & "\STOCK\"
  Dir1.Path = RutaBackup
  File1.FileName = Dir1.Path & "\*.*"
  
  FStockInventario.Caption = "INGRESO DE STOCK POR ARCHIVO"
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY CI_RUC "
  SelectAdodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     AdoAux.Recordset.MoveLast
     Codigo = AdoAux.Recordset.Fields("CI_RUC")
  End If
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FStockInventario
  ConectarAdodc AdoAux
  ConectarAdodc AdoAct
  ConectarAdodc AdoInv
  ConectarAdodc AdoQuery
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

