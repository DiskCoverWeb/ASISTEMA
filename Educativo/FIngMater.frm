VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FIngMater 
   Caption         =   "Ingreso de Cuentas Contables"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin ComctlLib.TreeView TVCatalogo 
      Height          =   7050
      Left            =   105
      TabIndex        =   15
      ToolTipText     =   "Un click en el dibujo de la Cta. y presionar la tecla <DEL> Borra la Cta."
      Top             =   315
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12435
      _Version        =   327682
      Indentation     =   794
      Style           =   5
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TextConcepto 
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
      Left            =   6945
      MaxLength       =   45
      TabIndex        =   6
      Top             =   1920
      Width           =   3360
   End
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   315
      Top             =   2100
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
   Begin VB.Frame Frame2 
      Caption         =   "TIPO DE CUENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6945
      TabIndex        =   12
      Top             =   2280
      Width           =   3360
      Begin VB.OptionButton OpcG 
         Caption         =   "Grupo"
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
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OpcD 
         Caption         =   "Detalle"
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
         Left            =   210
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin MSMask.MaskEdBox MBoxCta 
      Height          =   330
      Left            =   6930
      TabIndex        =   2
      Top             =   405
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "C.CC.CC.CCC"
      Mask            =   "C.CC.CC.CCC"
      PromptChar      =   "_"
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
      Left            =   10605
      Picture         =   "FIngMater.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
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
      Picture         =   "FIngMater.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   210
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoDetalle 
      Height          =   330
      Left            =   315
      Top             =   2415
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
      Caption         =   "Detalle"
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
   Begin MSAdodcLib.Adodc AdoCtas 
      Height          =   330
      Left            =   315
      Top             =   2730
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
      Caption         =   "Ctas"
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
      Left            =   315
      Top             =   3045
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
   Begin MSAdodcLib.Adodc AdoPresupuestos 
      Height          =   330
      Left            =   315
      Top             =   1785
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
      Caption         =   "Presupuestos"
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
   Begin MSDataGridLib.DataGrid DGDetalle 
      Bindings        =   "FIngMater.frx":0D0C
      Height          =   4215
      Left            =   7560
      TabIndex        =   16
      Top             =   3045
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7435
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
   Begin VB.CommandButton Command3 
      Caption         =   "&Insertar"
      Height          =   492
      Left            =   5760
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   972
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7065
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   34
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngMater.frx":0D25
            Key             =   "UNO"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngMater.frx":117F
            Key             =   "DOS"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngMater.frx":1499
            Key             =   "TRES"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FIngMater.frx":17B3
            Key             =   "CUATRO"
         EndProperty
      EndProperty
   End
   Begin VB.Label LabelTipoCta 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6930
      TabIndex        =   13
      Top             =   1140
      Width           =   1800
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NIVEL"
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
      Left            =   6930
      TabIndex        =   14
      Top             =   825
      Width           =   1800
   End
   Begin VB.Label LabelCtaSup 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   8745
      TabIndex        =   4
      Top             =   405
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ELIJA LA CUENTA SI DESEA CAMBIAR DATOS"
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
      Width           =   6735
   End
   Begin VB.Label LabelNumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   8745
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta Superior"
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
      Left            =   8745
      TabIndex        =   3
      Top             =   90
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DETALLE"
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
      Left            =   6945
      TabIndex        =   5
      Top             =   1560
      Width           =   3360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CODIGO"
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
      Left            =   6930
      TabIndex        =   1
      Top             =   90
      Width           =   1800
   End
End
Attribute VB_Name = "FIngMater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cta_Ini As String
Dim Cta_Fin As String
Dim nodX As Node

Private Sub Command1_Click()
  If Nuevo Then GrabarCta (True) Else GrabarCta (False)
End Sub

Private Sub Command2_Click()
  Unload FIngMater
End Sub

Private Sub Form_Activate()
  DGDetalle.Visible = False
'  SQL1 = "SELECT * " _
'       & "FROM Trans_Presupuestos " _
'       & "WHERE Item = '" & NumEmpresa & "' "
'  SelectAdodc AdoPresupuestos, SQL1
'
'  SQL1 = "UPDATE Catalogo_SubCtas " _
'       & "SET Presupuesto = 0 " _
'       & "WHERE TC = 'G' " _
'       & "AND Item = '" & NumEmpresa & "' "
'  ConectarAdoExecute SQL1
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Estudiantil " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoCta, sSQL
  FormatoMaskMat MBoxCta
  RatonReloj
  With AdoCta.Recordset
   If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
           If Len(.Fields("Codigo")) = 1 Then
              Codigo = "E" & .Fields("Codigo")
              Cta_Sup = .Fields("Codigo")
              Cuenta = .Fields("Codigo") & " - " & .Fields("Detalle")
              AddNewCta .Fields("TC")
           Else
              Codigo = "E" & .Fields("Codigo")
              Cta_Sup = "E" & CodigoCuentaSup(.Fields("Codigo"))
              Cuenta = .Fields("Codigo") & " - " & .Fields("Detalle")
              If .Fields("TC") = "G" Then
                  AddNewCta "DG"
              Else
                  AddNewCta .Fields("TC")
              End If
           End If
          .MoveNext
        Loop
    End If
   End With
   RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCta
  ConectarAdodc AdoCtas
  ConectarAdodc AdoDetalle
  ConectarAdodc AdoPresupuestos
End Sub

Private Sub MBoxCta_GotFocus()
  MarcarTexto MBoxCta
End Sub

Private Sub MBoxCta_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBoxCta_LostFocus()
  Codigo = CodigoCuentaSup(CambioCodigoCta(MBoxCta.Text))
  If Codigo = "0" Then Codigo = CambioCodigoCta(MBoxCta.Text)
  sSQL = "SELECT Codigo " _
       & "FROM Catalogo_Estudiantil " _
       & "WHERE Codigo = '" & Codigo & "' " _
       & "AND Item = '" & NumEmpresa & "' "
  SelectAdodc AdoCtas, sSQL, False
  If (AdoCtas.Recordset.RecordCount <= 0) And (Len(Codigo) > 1) Then
     Cadena = "Warnign: No puede crear este Código," & vbCrLf _
            & "no existe Cuenta Superior "
     MsgBox Cadena
     MBoxCta.SetFocus
  Else
     LabelCtaSup.Caption = CambioCodigoCtaSup(CambioCodigoCta(MBoxCta.Text))
     Codigos = CambioCodigoCta(MBoxCta.Text)
     sSQL = "SELECT Codigo FROM Catalogo_Estudiantil " _
          & "WHERE Codigo = '" & Codigos & "' " _
          & "AND Item = '" & NumEmpresa & "' "
     SelectAdodc AdoCtas, sSQL
     If (AdoCtas.Recordset.RecordCount > 0) And (Nuevo) Then
        MsgBox "Esta Cuenta ya existe, vuelva a ingresar otra cuenta."
        MBoxCta.SetFocus
     Else
        LabelTipoCta.Caption = TiposCtaStrg(Codigo)
     End If
  End If
End Sub

Private Sub TextConcepto_GotFocus()
  MarcarTexto TextConcepto
End Sub

Private Sub TextConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextConcepto_LostFocus()
  TextoValido TextConcepto
End Sub

Public Sub LlenarCta()
  DGDetalle.Visible = False
  With AdoCta.Recordset
   If .RecordCount > 0 Then
      MBoxCta.Text = FormatoCodigoMat(.Fields("Codigo"))
      LabelCtaSup.Caption = CodigoCuentaSup(.Fields("Codigo"))
      TextConcepto.Text = .Fields("Detalle")
      Cadena = MBoxCta.Text
       Select Case .Fields("TC")
         Case "C": LabelTipoCta.Caption = "CICLO"
         Case "N": LabelTipoCta.Caption = "NIVEL"
         Case "P"
             LabelTipoCta.Caption = "PARALELO"
             sSQL = "SELECT Codigo,Cliente " _
                  & "FROM Clientes " _
                  & "WHERE Grupo = '" & Detalle & "' & '" & Cadena & "' "
             SelectDataGrid DGDetalle, AdoDetalle, sSQL
             DGDetalle.Visible = True
         Case "M":    LabelTipoCta.Caption = "MATERIA"
            sSQL = "SELECT Codigo,Cliente " _
                  & "FROM Clientes " _
                  & "WHERE Grupo = 'P' & '" & Cadena & "' "
             SelectDataGrid DGDetalle, AdoDetalle, sSQL
             DGDetalle.Visible = True
       End Select
      Nuevo = False
   Else
      Nuevo = True
   End If
   End With
   Label6.Visible = True
End Sub

Public Sub GrabarCta(NuevaCta As Boolean)
  Select Case Len(Codigo)
     Case 11: TipoCta = "M"
     Case 1: TipoCta = "C"
     Case 4: TipoCta = "N"
     Case 7: TipoCta = "P"
  End Select
  If OpcG.Value Then
     TextoValido TextConcepto, , True
  Else
     TextoValido TextConcepto
  End If
  If LabelCtaSup.Caption = "" Then LabelCtaSup.Caption = "0"
  Numero = 0
  Codigo1 = CambioCodigoCta(MBoxCta.Text)
  Codigo = "E" & Codigo1
  Cta_Sup = "E" & CodigoCuentaSup(Codigo1)
  Cuenta = Codigo1 & " - " & TextConcepto.Text
  Mensajes = "Esta seguro de Grabar la cuenta" & vbCrLf _
           & "No. [" & Codigo1 & "] - " & TextConcepto.Text
  Titulo = "Pregunta de grabación"
  If BoxMensaje = vbYes Then
     With AdoCta.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         .Find ("Codigo like '" & Codigo1 & "' ")
          If Not .EOF Then
'             Numero = .Fields("Clave")
             If OpcD.Value And Numero = 0 Then
                Numero = ReadSetDataNum("Numero Cuenta", True, True)
             End If
          Else
            .AddNew
            .Fields("Codigo") = Codigo1
             If OpcD.Value Then
                Numero = ReadSetDataNum("Numero Cuenta", True, True)
             End If
             AddNewCta TipoCta
          End If
      Else
         .AddNew
         .Fields("Codigo") = Codigo1
          If OpcD.Value Then
             Numero = ReadSetDataNum("Numero Cuenta", True, True)
          End If
          If OpcG.Value Then AddNewCta "DG" Else AddNewCta TipoCta
      End If
'     .Fields("Clave") = Numero
     .Fields("TC") = TipoCta
     .Fields("Detalle") = TextConcepto.Text
     .Fields("Item") = NumEmpresa
     .Fields("Lleno") = 0
     .Fields("Estado") = 1
     .Update
      UpdateCta TipoCta
     End With
  End If
  sSQL = "SELECT * " _
       & "FROM Catalogo_Estudiantil " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoCta, sSQL
  TVCatalogo.Refresh
  Label6.Visible = True
  Nuevo = False
End Sub

Public Sub NuevaCta()
  OpcNor.Value = True
  LabelNumero.Caption = "0"
  LabelNumero.Caption = ""
  TextConcepto.Text = ""
  TextPresupuesto.Text = ""
  LabelCtaSup.Caption = ""
  MBoxCta.Text = LimpiarMater
  Nuevo = True
  MBoxCta.SetFocus
End Sub

Private Sub TVCatalogo_DblClick()
  SiguienteControl
End Sub

Private Sub TVCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyI And CtrlDown Then
      Cta_Ini = SinEspaciosIzq(TVCatalogo.SelectedItem)
      Select Case Len(Cta_Ini)
          Case 4:
          Case 7
              Cadena2 = SinEspaciosDer(TVCatalogo.SelectedItem)
              If Len(Cadena2) > 3 Then Cadena2 = Mid(Cadena2, Len(Cadena2) - 3, 3) Else FALTANTE = 3 - Len(Cadena2)
              'MsgBox Cadena2 & vbCrLf & FALTANTE
              Cadena1 = ""
              Cadena1 = SinEspaciosIzq(TVCatalogo.SelectedItem)
              Cadena1 = QuitarPuntos(Cadena1)
              MsgBox Cadena1
        
      End Select
  End If
  If KeyCode = vbKeyU And CtrlDown Then
      Cta_Fin = SinEspaciosIzq(TVCatalogo.SelectedItem)
  End If
  If KeyCode = vbKeyDelete Then EliminarCta
End Sub

Private Sub TVCatalogo_LostFocus()
  Cadena = SinEspaciosIzq(TVCatalogo.SelectedItem)
  With AdoCta.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo like '" & Cadena & "' ")
       If Not .EOF Then LlenarCta
   End If
  End With
End Sub

Public Sub AddNewCta(TipoTC As String)
  If Len(Codigo) = 2 Then
     Set nodX = TVCatalogo.Nodes.Add(, , Codigo, Cuenta)
     nodX.Image = ImageList1.ListImages(1).key
     nodX.SelectedImage = ImageList1.ListImages(1).key
  Else
     Set nodX = TVCatalogo.Nodes.Add(Cta_Sup, tvwChild, Codigo, Cuenta)
     Select Case Len(Codigo)
       Case 5: TipoTC = "N"
       Case 8: TipoTC = "P"
       Case 12: TipoTC = "M"
     End Select
     Select Case TipoTC
       Case "C": IE = 1
       Case "N": IE = 2
       Case "P": IE = 3
       Case "M": IE = 4
     End Select
     nodX.Image = ImageList1.ListImages(IE).key
     nodX.SelectedImage = ImageList1.ListImages(IE).key
     TVCatalogo.Tag = Codigo
  End If
End Sub

Public Sub UpdateCta(TipoTC As String)
 ' TVCatalogo.SelectedItem = Cuenta
  Select Case Len(Codigo)
       Case 5: TipoTC = "N"
       Case 8: TipoTC = "P"
       Case 12: TipoTC = "M"
     End Select
  Select Case TipoTC
    Case "C": IE = 1
    Case "N": IE = 2
    Case "P": IE = 3
    Case "M": IE = 4
  End Select
  nodX.Image = ImageList1.ListImages(IE).key
  nodX.SelectedImage = ImageList1.ListImages(IE).key
End Sub

Public Sub EliminarCta()
  Codigo1 = CambioCodigoCta(MBoxCta.Text)
  Cadena = SinEspaciosIzq(TVCatalogo.SelectedItem)
  With AdoCta.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo like '" & Cadena & "' ")
       If Not .EOF Then
          sSQL = "SELECT Cta " _
               & "FROM Transacciones " _
               & "WHERE Cta = '" & Cadena & "' " _
               & "AND Item = '" & NumEmpresa & "' "
          SelectAdodc AdoCtas, sSQL, False
          If AdoCtas.Recordset.RecordCount > 0 Then
             Mensajes = "No se puede eliminar esta Cuenta," & vbCrLf _
                      & "porque tiene cuentas procesables."
             MsgBox Mensajes
          Else
             Mensajes = "Esta seguro que desea eliminar la " & vbCrLf _
                      & "Cuenta No. [" & Cadena & "]"
             Titulo = "Pregunta de Eliminacion"
             If BoxMensaje = vbYes Then
               .Delete
                TVCatalogo.Nodes.Remove TVCatalogo.SelectedItem.Index
'''                sSQL = "DELETE * FROM Catalogo_Estudiantil " _
'''                     & "WHERE Codigo = '" & Cadena & "' " _
'''                     & "AND Item = '" & NumEmpresa & "' "
'''                ConectarAdoExecute sSQL
             End If
          End If
       End If
   End If
  End With
End Sub

