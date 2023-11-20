VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form FCambioPedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CAMBIO DE VALORES DE PEDIDOS"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGAux1 
      Bindings        =   "FPedCamb.frx":0000
      Height          =   3270
      Left            =   105
      TabIndex        =   7
      Top             =   525
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5768
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cambiar de Orden"
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
      Left            =   10605
      Picture         =   "FPedCamb.frx":0016
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1995
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cambiar Rubro de Orden"
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
      Left            =   10605
      Picture         =   "FPedCamb.frx":06AC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1050
      Width           =   1380
   End
   Begin MSDataGridLib.DataGrid DGProducto 
      Bindings        =   "FPedCamb.frx":0D42
      Height          =   3270
      Left            =   2625
      TabIndex        =   2
      Top             =   525
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   5768
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
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
   Begin MSAdodcLib.Adodc AdoProducto 
      Height          =   330
      Left            =   735
      Top             =   1050
      Visible         =   0   'False
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
      Caption         =   "Producto"
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
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Grabar Cambios"
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
      Left            =   10605
      Picture         =   "FPedCamb.frx":0D5C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Salir"
      DisabledPicture =   "FPedCamb.frx":119E
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
      Left            =   10605
      MouseIcon       =   "FPedCamb.frx":1BE8
      Picture         =   "FPedCamb.frx":2632
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2940
      Width           =   1380
   End
   Begin VB.TextBox TxtOrden 
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
      Left            =   1365
      TabIndex        =   1
      Top             =   105
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   735
      Top             =   1365
      Visible         =   0   'False
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
      Caption         =   "AdoAux"
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
   Begin MSAdodcLib.Adodc AdoAux1 
      Height          =   330
      Left            =   735
      Top             =   1680
      Visible         =   0   'False
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
      Caption         =   "AdoAux"
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
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &ORDEN No."
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
End
Attribute VB_Name = "FCambioPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Listar_Pedidos()
  If IsNumeric(TxtOrden) Then
     sSQL = "SELECT Fecha,Producto,Cantidad,Orden_No,Opc1 As Agua,Opc2 As Seco,Opc3 As Lavado,Hora,Codigo,ID,Item "
  Else
     sSQL = "SELECT Fecha,Producto,Cantidad,Precio,Total,Orden_No,No_Hab,Total_IVA,Hora,Codigo,ID,Item "
  End If
  sSQL = sSQL & "FROM Trans_Pedidos " _
       & "WHERE Item = '" & NumEmpresa & "' "
  If IsNumeric(TxtOrden) Then
     sSQL = sSQL & "AND Orden_No = " & Orden_No & " " _
          & "ORDER BY Producto "
  Else
     sSQL = sSQL & "AND No_Hab = '" & Habitacion_No & "' " _
          & "ORDER BY Codigo "
  End If
  Select_Adodc_Grid DGProducto, AdoProducto, sSQL
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
  With AdoProducto.Recordset
   If .RecordCount > 0 Then
       Codigo = .Fields("No_Hab")
       CodigoInv = .Fields("Codigo")
       Producto = .Fields("Producto")
       Cadena = "ORDEN No. [" & Habitacion_No & "] - " _
              & CodigoInv & vbCrLf & vbCrLf _
              & Space(20) & Producto
       Habitacion_No = UCaseStrg(InputBox(Cadena, "CAMBIAR POR", ""))
       If Len(Habitacion_No) > 1 Then
          sSQL = "UPDATE Trans_Pedidos " _
               & "SET No_Hab = '" & Habitacion_No & "' " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND No_Hab = '" & Codigo & "' " _
               & "AND Codigo = '" & CodigoInv & "' "
          Ejecutar_SQL_SP sSQL
       End If
       Habitacion_No = Codigo
   End If
  End With
  Listar_Pedidos
End Sub

Private Sub Command3_Click()
  With AdoProducto.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       RatonReloj
       Do While Not .EOF
          Total_IVA = 0
          If .Fields("Total_IVA") > 0 Then Total_IVA = .Fields("Precio") * .Fields("Cantidad") * Porc_IVA
         .Fields("Total_IVA") = Total_IVA
         .Fields("Total") = .Fields("Precio") * .Fields("Cantidad")
         .Update
         .MoveNext
       Loop
   End If
  End With
  RatonNormal
  Unload Me
End Sub

Private Sub Command4_Click()
  With AdoProducto.Recordset
   If .RecordCount > 0 Then
       CodigoInv = .Fields("Codigo")
       Producto = .Fields("Producto")
       If IsNumeric(TxtOrden) Then
          Cadena = "ORDEN No. " & Orden_No & vbCrLf
          Nota_No = .Fields("Orden_No")
          Orden_No = Val(InputBox(Cadena, "CAMBIAR POR", ""))
          sSQL = "UPDATE Trans_Pedidos " _
               & "SET Orden_No = " & Orden_No & " " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Orden_No = " & Nota_No & " "
          Orden_No = Nota_No
       Else
          Cadena = "ORDEN No. " & Habitacion_No & vbCrLf
          Codigo = .Fields("No_Hab")
          Habitacion_No = UCaseStrg(InputBox(Cadena, "CAMBIAR POR", ""))
          If Habitacion_No = "" Then Habitacion_No = Ninguno
          sSQL = "UPDATE Trans_Pedidos " _
               & "SET No_Hab = '" & Habitacion_No & "' " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND No_Hab = '" & Codigo & "' "
          Habitacion_No = Codigo
       End If
       sSQL = sSQL & "AND Codigo = '" & CodigoInv & "' "
       Ejecutar_SQL_SP sSQL
   End If
  End With
  Listar_Pedidos
End Sub

Private Sub Form_Activate()
  Habitacion_No = Ninguno
  Listar_Pedidos
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FCambioPedidos
  ConectarAdodc AdoAux
  ConectarAdodc AdoAux1
  ConectarAdodc AdoProducto
End Sub

Private Sub TxtOrden_GotFocus()
  TxtOrden.Text = ""
End Sub

Private Sub TxtOrden_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtOrden_LostFocus()
  If IsNumeric(TxtOrden) Then
     Orden_No = Val(TxtOrden)
     Command2.Enabled = False
  Else
     Habitacion_No = UCaseStrg(TxtOrden)
     Command2.Enabled = True
  End If
  Listar_Pedidos
  sSQL = "SELECT No_Hab,SUM(Total) As Consumo " _
       & "FROM Trans_Pedidos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "GROUP BY No_Hab " _
       & "ORDER BY No_Hab "
  Select_Adodc_Grid DGAux1, AdoAux1, sSQL
  DGAux1.Caption = "C O N S U M O S"
End Sub
