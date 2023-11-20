VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FCambioFacturas 
   Caption         =   "NUEVA AUTORIZACION"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataList DLFacturas 
      Bindings        =   "CambioFacturas.frx":0000
      DataSource      =   "AdoFacturas"
      Height          =   1020
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1799
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2520
      TabIndex        =   1
      Top             =   735
      Width           =   960
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
      Height          =   540
      Left            =   2520
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoFacturas 
      Height          =   330
      Left            =   105
      Top             =   2205
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
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
   Begin VB.Label LblRango 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      TabIndex        =   2
      Top             =   1155
      Width           =   2325
   End
End
Attribute VB_Name = "FCambioFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   Mensajes = "Cambiar: " & LblRango.Caption & vbCrLf _
             & "Por la Autorizacion: " & DLFacturas.Text
   Titulo = "Formulario de Grabación."
   If BoxMensaje = vbYes Then
      sSQL = "UPDATE Facturas " _
           & "SET Autorizacion = '" & DLFacturas.Text & "' " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TC = '" & FA.TC & "' " _
           & "AND Serie = '" & FA.Serie & "' " _
           & "AND Autorizacion = '" & FA.Autorizacion & "' " _
           & "AND Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " "
      ConectarAdoExecute sSQL
      sSQL = "UPDATE Detalle_Factura " _
           & "SET Autorizacion = '" & DLFacturas.Text & "' " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TC = '" & FA.TC & "' " _
           & "AND Serie = '" & FA.Serie & "' " _
           & "AND Autorizacion = '" & FA.Autorizacion & "' " _
           & "AND Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " "
      ConectarAdoExecute sSQL
      sSQL = "UPDATE Trans_Abonos " _
           & "SET Autorizacion = '" & DLFacturas.Text & "' " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND TP = '" & FA.TC & "' " _
           & "AND Serie = '" & FA.Serie & "' " _
           & "AND Autorizacion = '" & FA.Autorizacion & "' " _
           & "AND Factura BETWEEN " & Factura_Desde & " and " & Factura_Hasta & " "
      ConectarAdoExecute sSQL
      Unload FCambioFacturas
  End If
End Sub

Private Sub Command2_Click()
  Unload FCambioFacturas
End Sub

Private Sub Form_Activate()
  sSQL = "SELECT Autorizacion " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TL <> " & Val(adFalse) & " " _
       & "AND Autorizacion <> '" & FA.Autorizacion & "' " _
       & "GROUP BY Autorizacion " _
       & "ORDER BY Autorizacion "
  SelectDBList DLFacturas, AdoFacturas, sSQL, "Autorizacion"
  LblRango.Caption = "Auto.: " & FA.Autorizacion & vbCrLf _
                   & "Serie: " & FA.TC & "-" & FA.Serie & vbCrLf _
                   & "Desde: " & Factura_Desde & vbCrLf _
                   & "Hasta: " & Factura_Hasta
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FCambioFacturas
  ConectarAdodc AdoFacturas
End Sub
