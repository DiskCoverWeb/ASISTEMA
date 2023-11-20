VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form FListDiario 
   Caption         =   "LIBRO DIARIO"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anular"
            Object.ToolTipText     =   "Anular Comprobante"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Comprobante"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Primero"
            Object.ToolTipText     =   "Primer Comprobante"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Anterior"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Siguiente"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Ultimo"
            Object.ToolTipText     =   "Ultimo Comprobante"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSDBGrid.DBGrid DBGDiario 
      Bindings        =   "ListDiar.frx":0000
      Height          =   4740
      Left            =   105
      OleObjectBlob   =   "ListDiar.frx":0014
      TabIndex        =   7
      Top             =   1680
      Width           =   11355
   End
   Begin VB.Data DataRet 
      Caption         =   "Ret"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5250
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data DataAux 
      Caption         =   "Aux"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4725
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5985
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
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
      Height          =   435
      Left            =   4200
      TabIndex        =   17
      Top             =   6510
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Data DataSubCtas 
      Caption         =   "SubCtas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2415
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5985
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reingresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2835
      TabIndex        =   13
      Top             =   6510
      Width           =   1275
   End
   Begin VB.Data DataTrans 
      Caption         =   "Trans"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5985
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Data DataComprobantes 
      Caption         =   "Comprobantes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5670
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   5
      Top             =   6510
      Width           =   1275
   End
   Begin VB.CommandButton CmdCancelar 
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
      Height          =   435
      Left            =   1470
      TabIndex        =   6
      Top             =   6510
      Width           =   1275
   End
   Begin MSDBCtls.DBCombo DBCComp 
      Bindings        =   "ListDiar.frx":09E4
      DataSource      =   "DataComp"
      Height          =   360
      Left            =   1470
      TabIndex        =   9
      Top             =   525
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   635
      _Version        =   327680
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data DataComp 
      Caption         =   "Comp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2415
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5670
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label3 
      Caption         =   " Recibi de:"
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
      Left            =   3045
      TabIndex        =   15
      Top             =   630
      Width           =   960
   End
   Begin VB.Label LabelRecibi 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recibi"
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
      Left            =   4095
      TabIndex        =   14
      Top             =   525
      Width           =   4110
   End
   Begin VB.Label Label4 
      Caption         =   " Concepto: "
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
      Left            =   210
      TabIndex        =   12
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label LabelConcepto 
      BackColor       =   &H80000009&
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
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1470
      TabIndex        =   11
      Top             =   945
      Width           =   6750
   End
   Begin VB.Label LabelEst 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9240
      TabIndex        =   10
      Top             =   945
      Width           =   1800
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListDiar.frx":09F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListDiar.frx":0B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListDiar.frx":0CAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListDiar.frx":11C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListDiar.frx":16D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ListDiar.frx":1BE5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label LabelFecha 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9240
      TabIndex        =   8
      Top             =   525
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   " Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8295
      TabIndex        =   1
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   " Diario No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   525
      Width           =   1170
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Totales"
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
      Left            =   5985
      TabIndex        =   2
      Top             =   6510
      Width           =   1065
   End
   Begin VB.Label LabelDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   7140
      TabIndex        =   3
      Top             =   6510
      Width           =   1905
   End
   Begin VB.Label LabelHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9135
      TabIndex        =   4
      Top             =   6510
      Width           =   1905
   End
End
Attribute VB_Name = "FListDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancelar_Click()
   Unload FListDiario
End Sub

Private Sub CmdGrabar_Click()
  NumItem = NumEmpresa
  NumComp = Val(DBCComp.Text)
  If OpcCoop Then
     NumItem = SinEspaciosIzq(DBCComp.Text)
     NumComp = SinEspaciosDer(DBCComp.Text)
  End If
  ImprimirComprobantesDe False, CompDiario, NumComp, NumItem, DataComprobantes, DataTrans, , DataRet
  ListarDiario NumItem, NumComp
  DBCComp.SetFocus
End Sub

Private Sub Command1_Click()
   Mensajes = "Seguro que quiere Reingresar el Comprobante No. " & NumComp
   Titulo = "Pregunta de Reingreso"
   If BoxMensaje = 6 Then
      NumItem = NumEmpresa
      If OpcCoop Then NumItem = SinEspaciosIzq(DBCComp.Text)
      NumCompDia = NumComp
      Cta_General = Ninguno
      Unload FListDiario
      FDiario.Show
   End If
End Sub

Private Sub Command2_Click()
   Mensajes = "Seguro que quiere Modificar el Comprobante No. " & NumComp
   Titulo = "Pregunta de Reingreso"
   If BoxMensaje = 6 Then
      NumItem = NumEmpresa
      If OpcCoop Then NumItem = SinEspaciosIzq(DBCComp.Text)
      NumCompDia = NumComp
      Cta_General = Ninguno
      'ModificarComprobantes
      Unload FListDiario
      'FDiario.Show
   End If
End Sub

Private Sub DBCComp_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      NumItem = NumEmpresa
      NumComp = Val(DBCComp.Text)
      If OpcCoop Then
         NumItem = SinEspaciosIzq(DBCComp.Text)
         NumComp = SinEspaciosDer(DBCComp.Text)
      End If
      ListarDiario NumItem, NumComp
   End If
End Sub

Private Sub Form_Activate()
  Command1.Enabled = Supervisor
  If OpcCoop Then
     sSQL = "SELECT (Item & ' ' & Numero) As Numeros FROM Comprobantes "
     sSQL = sSQL & "WHERE TP = '" & CompDiario & "' "
     sSQL = sSQL & "ORDER BY Item,Numero "
     SelectDBCombo DBCComp, DataComp, sSQL, "Numeros", True
  Else
     sSQL = "SELECT Numero FROM Comprobantes "
     sSQL = sSQL & "WHERE TP = '" & CompDiario & "' "
     sSQL = sSQL & "ORDER BY Numero "
     SelectDBCombo DBCComp, DataComp, sSQL, "Numero", True
  End If
  If DataComp.Recordset.RecordCount > 0 Then
     NumItem = NumEmpresa
     NumComp = Val(DBCComp.Text)
     If OpcCoop Then
        NumItem = SinEspaciosIzq(DBCComp.Text)
        NumComp = SinEspaciosDer(DBCComp.Text)
     End If
     ListarDiario NumItem, NumComp
     RatonNormal
     DBCComp.SetFocus
  Else
     Unload FListDiario
     RatonNormal
  End If
End Sub

Private Sub Form_Load()
   CentrarForm FListDiario
  'Abriendo bases relacionadas
   DataAux.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataRet.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataComp.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataTrans.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataComprobantes.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
   DataSubCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  With DataComp.Recordset
  Select Case Button.Key
    Case "Anular"
        NumItem = NumEmpresa
        NumComp = Val(DBCComp.Text)
        If OpcCoop Then
           NumItem = SinEspaciosIzq(DBCComp.Text)
           NumComp = SinEspaciosDer(DBCComp.Text)
        End If
        Mensajes = "Seguro que quiere anular el Comprobante No. " & NumComp
        Titulo = "Pregunta de Anulacion"
        TipoDeCaja = 4 + 32: ResultBox = MsgBox(Mensajes, TipoDeCaja, Titulo)
        If ResultBox = 6 Then
           sSQL = "UPDATE Comprobantes SET T = '" & Anulado & "', "
           sSQL = sSQL & "Concepto = 'C O M P R O B A N T E   A N U L A D O' "
           sSQL = sSQL & "WHERE Numero = " & NumComp & " "
           sSQL = sSQL & "AND TP = '" & CompDiario & "' "
           sSQL = sSQL & "AND Item = " & NumItem & " "
           DataComprobantes.Database.Execute sSQL
          'Actualizar transacciones
           sSQL = "UPDATE Transacciones SET T = '" & Anulado & "', "
           sSQL = sSQL & "Debe = 0,Haber = 0,Saldo = 0 "
           sSQL = sSQL & "WHERE TP = '" & CompDiario & "' "
           sSQL = sSQL & "AND Numero = " & NumComp & " "
           sSQL = sSQL & "AND Item = " & NumItem & " "
           DataTrans.Database.Execute sSQL
          'Actualizar transacciones
           sSQL = "UPDATE TransaccionesSC SET T = '" & Anulado & "', "
           sSQL = sSQL & "Debitos = 0,Creditos = 0,Saldo = 0 "
           sSQL = sSQL & "WHERE TP = '" & CompDiario & "' "
           sSQL = sSQL & "AND Numero = " & NumComp & " "
           sSQL = sSQL & "AND Item = " & NumItem & " "
           DataTrans.Database.Execute sSQL
          'Actualizar kardex
           sSQL = "DELETE * FROM Kardex "
           sSQL = sSQL & "WHERE TP = '" & CompDiario & "' "
           sSQL = sSQL & "AND Numero = " & NumComp & " "
           DataTrans.Database.Execute sSQL
        End If
    Case "Imprimir"
        NumItem = NumEmpresa
        NumComp = Val(DBCComp.Text)
        If OpcCoop Then
           NumItem = SinEspaciosIzq(DBCComp.Text)
           NumComp = SinEspaciosDer(DBCComp.Text)
        End If
        ImprimirComprobantesDe False, CompDiario, NumComp, NumItem, DataComprobantes, DataTrans, , DataRet
        ListarDiario NumItem, NumComp
        DBCComp.SetFocus
    Case "Primero"
        .MoveFirst
    Case "Anterior"
        .MovePrevious
        If .BOF Then .MoveFirst
    Case "Siguiente"
        .MoveNext
        If .EOF Then .MoveLast
    Case "Ultimo"
        .MoveLast
  End Select
  If Button.Key <> "Imprimir" Then
     If OpcCoop Then DBCComp.Text = .Fields("Numeros") Else DBCComp.Text = .Fields("Numero")
     NumItem = NumEmpresa
     NumComp = Val(DBCComp.Text)
     If OpcCoop Then
        NumItem = SinEspaciosIzq(DBCComp.Text)
        NumComp = SinEspaciosDer(DBCComp.Text)
     End If
     ListarDiario NumItem, NumComp
  End If
  End With
End Sub

Public Sub ListarDiario(NoAgencia, Comp_No)
  RatonReloj
  sSQL = "SELECT * FROM Comprobantes "
  sSQL = sSQL & "WHERE Numero = " & Comp_No & " "
  sSQL = sSQL & "AND TP = '" & CompDiario & "' "
  sSQL = sSQL & "AND Item = " & NoAgencia & " "
  SelectData DataComprobantes, sSQL, False
  With DataComprobantes.Recordset
   If .RecordCount > 0 Then
       If .Fields("T") = Anulado Then
           LabelEst.Caption = "Anulado"
       Else
           LabelEst.Caption = "Normal"
       End If
       LabelFecha.Caption = .Fields("Fecha")
       LabelConcepto.Caption = .Fields("Concepto")
       LabelRecibi.Caption = .Fields("Beneficiario")
   Else
       MsgBox "El Comprobante no exite."
   End If
  End With
  sSQL = "SELECT ID FROM Transacciones "
  sSQL = sSQL & "WHERE TP = '" & CompIngreso & "' "
  sSQL = sSQL & "AND Numero = " & Comp_No & " "
  sSQL = sSQL & "ORDER BY ID "
  SelectData DataTrans, sSQL, False
  If DataTrans.Recordset.RecordCount > 0 Then
     ID_Trans = DataTrans.Recordset.Fields("ID")
     NumCompDia = Comp_No
  End If
  If OpcCoop Then
     sSQL = "SELECT Cta,Ca.Cuenta,Debe,Haber,Debe_ME,Haber_ME "
  Else
     sSQL = "SELECT Cta,Ca.Cuenta,(Debe_ME-Haber_ME) As Parcial_ME,Debe,Haber "
  End If
  sSQL = sSQL & "FROM Transacciones,Catalogo As Ca "
  sSQL = sSQL & "WHERE TP = '" & CompDiario & "' "
  sSQL = sSQL & "AND Numero = " & Comp_No & " "
  sSQL = sSQL & "AND Item = " & NoAgencia & " "
  sSQL = sSQL & "AND Ca.Codigo = Cta "
  sSQL = sSQL & "ORDER BY Debe DESC,Debe_ME DESC,Cta "
  SelectDBGrid DBGDiario, DataTrans, sSQL
  SumaTotalAsientos DataTrans
  LabelDebe.Caption = Format(SumaDebe, "#,##0.00")
  LabelHaber.Caption = Format(SumaHaber, "#,##0.00")
  RatonNormal
End Sub
