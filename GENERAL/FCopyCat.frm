VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FCopyCat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COPIAR CATALOGO DE OTRAS EMPRESAS"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CheqCatalogo 
      Caption         =   "Catalogo de Cuentas"
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
      Left            =   7455
      TabIndex        =   2
      Top             =   315
      Width           =   3270
   End
   Begin VB.CheckBox CheqSubCP 
      Caption         =   "SubCuentas de CxC y CxP"
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
      Left            =   7455
      TabIndex        =   6
      Top             =   2415
      Width           =   3270
   End
   Begin VB.CheckBox CheqSubCta 
      Caption         =   "SubCuentas de Ingreso y Gastos"
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
      Left            =   7455
      TabIndex        =   5
      Top             =   1890
      Width           =   3270
   End
   Begin VB.CheckBox CheqFact 
      Caption         =   "Seteos de Facturación"
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
      Left            =   7455
      TabIndex        =   4
      Top             =   1365
      Width           =   3270
   End
   Begin VB.CheckBox CheqSetImp 
      Caption         =   "Seteos de Impresión"
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
      Left            =   7455
      TabIndex        =   3
      Top             =   840
      Width           =   3270
   End
   Begin MSDataListLib.DataList DLEmpresa 
      Bindings        =   "FCopyCat.frx":0000
      DataSource      =   "AdoEmp"
      Height          =   2460
      Left            =   105
      TabIndex        =   1
      Top             =   315
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   4339
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
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
      Left            =   11025
      Picture         =   "FCopyCat.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   315
      Width           =   1170
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
      Height          =   855
      Left            =   11025
      Picture         =   "FCopyCat.frx":08DF
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1365
      Width           =   1170
   End
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   210
      Top             =   840
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Cta"
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
   Begin MSAdodcLib.Adodc AdoEmp 
      Height          =   330
      Left            =   210
      Top             =   1155
      Width           =   2430
      _ExtentX        =   4286
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
      Caption         =   "Emp"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   210
      Top             =   1470
      Width           =   2430
      _ExtentX        =   4286
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ELIJA LA EMPRESA A COPIAR EL CATALOGO"
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
      Width           =   7155
   End
End
Attribute VB_Name = "FCopyCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim PeriodoCopy As String
  PeriodoCopy = Periodo_Contable
  Cadena = DLEmpresa.Text
  NumItem = Ninguno
  If Cadena = "" Then Cadena = Ninguno
  With AdoEmp.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Empresa LIKE '" & Cadena & "' ")
       If Not .EOF Then NumItem = .Fields("Item")
   End If
  End With
  If NumItem <> Ninguno Then
     Mensajes = "Seguro de Copiar el Catalogo de:" & vbCrLf _
              & "(" & NumItem & ") " & Cadena & vbCrLf _
              & "Este proceso reemplazará el catalogo actual."
     Titulo = "Pregunta de Copia"
     If BoxMensaje = vbYes Then
        If Si_No Then
           PeriodoCopy = InputBox("Escriba el periodo a donde Copiar:", "COPIAR CATALOGOS", Periodo_Contable)
           If Not IsDate(PeriodoCopy) Then
              MsgBox "Usted no ha escrito el periodo correcto"
              PeriodoCopy = Periodo_Contable
              Unload FCopyCat
           End If
        End If
        If CheqCatalogo.value = 1 Then
           Copiar_Tabla_SP "Catalogo_Cuentas", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
           Copiar_Tabla_SP "Codigos", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
           Copiar_Tabla_SP "Ctas_Proceso", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
        End If
        If CheqFact.value = 1 Then
           Copiar_Tabla_SP "Catalogo_Lineas", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
           Copiar_Tabla_SP "Catalogo_Productos", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
        End If
        If CheqSubCta.value = 1 Then Copiar_Tabla_SP "Catalogo_SubCtas", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
        If CheqSubCP.value = 1 Then Copiar_Tabla_SP "Catalogo_CxCxP", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
        
        If CheqSetImp.value = 1 Then
           Copiar_Tabla_SP "Formato", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
           Copiar_Tabla_SP "Seteos_Documentos", NumItem, NumEmpresa, PeriodoCopy, Periodo_Contable
        End If
        RatonNormal
        MsgBox "Proceso terminado con éxito"
        Unload FCopyCat
     End If
  End If
  
End Sub

Private Sub Command2_Click()
  Unload FCopyCat
End Sub

Private Sub Form_Activate()
  RatonNormal
  sSQL = "SELECT Empresa,Item " _
       & "FROM Empresas "
  If Si_No Then
     sSQL = sSQL & "WHERE Item = '" & NumEmpresa & "' "
     Command1.Caption = "&Cual Periodo"
  Else
     sSQL = sSQL & "WHERE Item <> '" & NumEmpresa & "' "
     Command1.Caption = "&Aceptar"
  End If
  sSQL = sSQL & "ORDER BY Empresa,Item "
  SelectDB_List DLEmpresa, AdoEmp, sSQL, "Empresa"
  If AdoEmp.Recordset.RecordCount <= 0 Then
     MsgBox "No tiene empresas a quien copiar"
     Unload FCopyCat
  Else
     DLEmpresa.SetFocus
  End If
End Sub

Private Sub Form_Load()
  CentrarForm FCopyCat
  ConectarAdodc AdoAux
  ConectarAdodc AdoCta
  ConectarAdodc AdoEmp
End Sub
