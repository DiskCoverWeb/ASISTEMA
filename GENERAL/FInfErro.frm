VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FInfoError 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORMULARIO DE INFORME DE ERRORES"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13680
   Icon            =   "FInfErro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command3 
      Caption         =   "&Excel"
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
      Left            =   12705
      Picture         =   "FInfErro.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   945
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DGInfoError 
      Bindings        =   "FInfErro.frx":12D8
      Height          =   7155
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   12621
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777088
      ColumnHeaders   =   0   'False
      ForeColor       =   128
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "Courier New"
         Size            =   9
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
      Left            =   12705
      Picture         =   "FInfErro.frx":12F3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1785
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
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
      Left            =   12705
      Picture         =   "FInfErro.frx":1BBD
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   855
   End
   Begin MSAdodcLib.Adodc AdoInfoError 
      Height          =   330
      Left            =   105
      Top             =   7245
      Width           =   12510
      _ExtentX        =   22066
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
      Caption         =   "InfoError"
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
Attribute VB_Name = "FInfoError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Eliminar_Tabla_Temporal
    Unload Me
End Sub

Private Sub Command2_Click()
Dim IniX As Single
Dim IniY As Single
On Error GoTo Errorhandler

RatonReloj
DGInfoError.Visible = False
SQLMsg1 = "IMPRESION DE ERRORES"
SQLMsg2 = "Fecha de Error del Archivo: " & Mifecha
SQLMsg3 = "Fecha de Impresion: " & FechaSistema & " - Usuario: " & NombreUsuario
Mensajes = "Seguro de Imprimir en:" & vbCrLf & Printer.DeviceName & "?"
Titulo = "IMPRESION"
Bandera = False
SetPrinters.Show 1
If PonImpresoraDefecto(SetNombrePRN) Then
    Escala_Centimetro 1, TipoCourierNew, 8
    'Iniciamos la impresion
    Pagina = 1
    IniX = 1: IniY = 0.5
    Encabezado IniX, 19
    Printer.FontName = TipoCourierNew
    Printer.FontItalic = False
    With AdoInfoError.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            Printer.CurrentX = IniX
            Printer.CurrentY = PosLinea
            Printer.Print .fields("Texto")
            PosLinea = PosLinea + Printer.TextHeight("H") + 0.1
            If PosLinea >= LimiteAlto Then
               Printer.NewPage
               PosLinea = IniY + 2
            End If
           .MoveNext
        Loop
     End If
    End With
    'Producto = InsertarLinea
    Printer.EndDoc
    Eliminar_Tabla_Temporal
    RatonNormal
    Unload Me
    Exit Sub
Errorhandler:
    Eliminar_Tabla_Temporal
    RatonNormal
    ErrorDeImpresion
    Unload FInfoError
    Exit Sub
Else
    Eliminar_Tabla_Temporal
    RatonNormal
    Unload FInfoError
End If
End Sub

Private Sub Command3_Click()
    DGInfoError.Visible = False
    GenerarDataTexto FInfoError, AdoInfoError, True
    Eliminar_Tabla_Temporal
    RatonNormal
    Unload FInfoError
End Sub

Private Sub Form_Activate()
Dim cSQL As String
Dim WidthText As Single
Dim AnchoMax As Byte
    DGInfoError.Visible = False
    cSQL = "SELECT Texto " _
         & "FROM Tabla_Temporal " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Modulo = '" & NumModulo & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "ORDER BY ID "
    Select_Adodc_Grid DGInfoError, AdoInfoError, cSQL
    With AdoInfoError.Recordset
     If .RecordCount > 0 Then
         AnchoMax = AdoInfoError.Recordset.fields(0).DefinedSize
         WidthText = DGInfoError.Parent.TextWidth(String$(AnchoMax, "H"))
         DGInfoError.Columns(0).width = WidthText
         DGInfoError.Refresh
         DGInfoError.Visible = True
         Unload FEsperar
         FInfoError.WindowState = vbNormal
         RatonNormal
     Else
         RatonNormal
         Unload FEsperar
         Unload FInfoError
     End If
    End With
End Sub

Private Sub Form_Load()
    RatonReloj
    Imagen_Esperar "Iniciamos el informe de Errores"
    CentrarForm FInfoError
    ConectarAdodc AdoInfoError
End Sub


Public Sub Eliminar_Tabla_Temporal()
    sSQL = "DELETE * " _
         & "FROM Tabla_Temporal " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Modulo = '" & NumModulo & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' "
    Ejecutar_SQL_SP sSQL
End Sub
