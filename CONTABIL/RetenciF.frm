VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Retencion1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario de Retenciones"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "RetenciF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextPorc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   6825
      MaxLength       =   4
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   2730
      Width           =   960
   End
   Begin VB.TextBox TextValorR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   5145
      MaxLength       =   10
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   2730
      Width           =   1590
   End
   Begin VB.TextBox TextFactura 
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "0"
      Top             =   2730
      Width           =   1380
   End
   Begin VB.ListBox LstTipoComp 
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
      Left            =   105
      TabIndex        =   2
      Top             =   2415
      Width           =   1905
   End
   Begin VB.TextBox TxtCompRet 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2100
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "0"
      Top             =   2730
      Width           =   1485
   End
   Begin MSDataListLib.DataList DLTipoRet 
      Bindings        =   "RetenciF.frx":0442
      DataSource      =   "AdoTipoRet"
      Height          =   1635
      Left            =   105
      TabIndex        =   1
      Top             =   630
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2884
      _Version        =   393216
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGRet 
      Bindings        =   "RetenciF.frx":045B
      Height          =   1380
      Left            =   105
      TabIndex        =   13
      Top             =   3150
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2434
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      BackColor       =   16761024
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
            Format          =   "0%"
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
   Begin MSAdodcLib.Adodc AdoDetRet 
      Height          =   330
      Left            =   2100
      Top             =   3360
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
      Caption         =   "DetRet"
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
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Continuar"
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
      Left            =   105
      Picture         =   "RetenciF.frx":0473
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4620
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc AdoTipoRet 
      Height          =   330
      Left            =   2100
      Top             =   3780
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
      Caption         =   "TipoRet"
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
   Begin VB.Label LabelValorRet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   7875
      TabIndex        =   12
      Top             =   2730
      Width           =   1485
   End
   Begin VB.Label LabelCuentaR 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONCEPTO DE RETENCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   4725
      TabIndex        =   0
      Top             =   105
      Width           =   4635
   End
   Begin VB.Label LblCliente 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONCEPTO DE RETENCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   105
      TabIndex        =   17
      Top             =   105
      Width           =   4635
   End
   Begin VB.Label LabelTotalRet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   7875
      TabIndex        =   15
      Top             =   4620
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Retenido"
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
      Left            =   6405
      TabIndex        =   16
      Top             =   4620
      Width           =   1485
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Retenido"
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
      Left            =   7875
      TabIndex        =   11
      Top             =   2415
      Width           =   1485
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reten. %"
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
      Left            =   6825
      TabIndex        =   9
      Top             =   2415
      Width           =   960
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Factura"
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
      Left            =   5145
      TabIndex        =   7
      Top             =   2415
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factura No."
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
      TabIndex        =   5
      Top             =   2415
      Width           =   1380
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Retención No."
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
      Left            =   2100
      TabIndex        =   3
      Top             =   2415
      Width           =   1485
   End
End
Attribute VB_Name = "Retencion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
  Total = 0: Total_DetRet = 0
  With AdoDetRet.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          Total = Total + .Fields("VALOR_FACT")
          Total_DetRet = Total_DetRet + .Fields("VALOR_RET")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  'MsgBox Total_DetRet
  Unload Retencion1
End Sub

Private Sub DGRet_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar la Transaccion."
  Titulo = "Confirmación de eliminación"
  If BoxMensaje = vbYes Then Cancel = False Else Cancel = True
End Sub

Private Sub DLTipoRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Command2.SetFocus
End Sub

Private Sub Form_Activate()
Dim NumRet As Long
   If OpcDH = 2 Then
      'LabelCompRet.Caption = Format(ReadSetDataNum("Retencion", True, False), "000000")
      LstTipoComp.Clear
      LstTipoComp.AddItem "FACTURA"
      LstTipoComp.AddItem "NOTA DE VENTA"
      LstTipoComp.AddItem "LIQUIDACION"
      LstTipoComp.AddItem "TICKET"
      
      sSQL = "SELECT CODIGO & '  ' & TIPO_RET As TipoRet " _
           & "FROM Tipo_Reten " _
           & "ORDER BY CODIGO "
      SelectDBList DLTipoRet, AdoTipoRet, sSQL, "TipoRet"
      sSQL = "SELECT * FROM Asiento_R " _
           & "WHERE CTA = '" & Codigo & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND T_No = " & Trans_No & " " _
           & "AND CodigoU = '" & CodigoUsuario & "' "
      SelectDataGrid DGRet, AdoDetRet, sSQL
      LabelCuentaR.Caption = "Codigo: " & Codigo & vbCrLf & "Cuenta: " & Nombre_Cta_Ret
      LblCliente.Caption = CodigoCliente & vbCrLf & NombreCliente
      DGRet.Visible = True
      'LabelTotal.Caption = Format(ValorDH, "#,##0.00")
      DLTipoRet.SetFocus
      Total = 0
   Else
      Unload Retencion1
   End If
End Sub

Private Sub Form_Load()
   If OpcDH = 2 Then
      CentrarForm Retencion1
      ConectarAdodc AdoDetRet
      ConectarAdodc AdoTipoRet
   End If
End Sub

Private Sub LstTipoComp_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextFactura_GotFocus()
  TextFactura.Text = ""
End Sub

Private Sub TextFactura_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextFactura_LostFocus()
  TextoValido TextFactura, True
End Sub

Private Sub TextPorc_Change()
   Total_DetRet = 0
   If TextPorc.Text <> "" Then
      Total_DetRet = Round(CDbl(TextValorR.Text) * CDbl(TextPorc.Text) / 100, 2)
   End If
   LabelValorRet.Caption = Format(Total_DetRet, "#,##0.00")
End Sub

Private Sub TextPorc_GotFocus()
   TextPorc.Text = ""
End Sub

Private Sub TextPorc_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextPorc_LostFocus()
  If TextPorc.Text = "" Then TextPorc.Text = "0"
  TextoValido TextFactura, True
  'TextPorc.Text = Round(TextPorc.Text,2)
  If (0 < Round(TextPorc.Text, 2)) And (Round(TextPorc.Text, 2) <= 100) Then
     Si_No = False
     If Mid(Grupo_No, 1, 3) = "EMP" Then
        Mensajes = "Salario Neto"
        Titulo = "CONFIRMACION"
        If BoxMensaje = vbYes Then Si_No = True
     End If
     SetAddNew AdoDetRet
     SetFields AdoDetRet, "ME", Moneda_US
     SetFields AdoDetRet, "CTA", Codigo
     SetFields AdoDetRet, "FACTURA_No", TextFactura.Text
     SetFields AdoDetRet, "Retenc_No", TxtCompRet.Text
     SetFields AdoDetRet, "VALOR_FACT", Round(TextValorR.Text, 2)
     SetFields AdoDetRet, "PORC", Round(TextPorc.Text, 2)
     SetFields AdoDetRet, "VALOR_RET", Round(Total_DetRet, 2)
     SetFields AdoDetRet, "Item", NumEmpresa
     SetFields AdoDetRet, "T_No", Trans_No
     SetFields AdoDetRet, "Tipo_Ret", SinEspaciosIzq(DLTipoRet.Text)
     SetFields AdoDetRet, "TD", Mid(LstTipoComp.Text, 1, 1)
     SetFields AdoDetRet, "TC", SubCta
     SetFields AdoDetRet, "SN", Si_No
     SetFields AdoDetRet, "Codigo", CodigoCliente
     SetFields AdoDetRet, "CodigoU", CodigoUsuario
     SetUpdate AdoDetRet
  End If
  With AdoDetRet.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Total = 0: Total_DetRet = 0
       Do While Not .EOF
          Total = Total + .Fields("VALOR_FACT")
          Total_DetRet = Total_DetRet + .Fields("VALOR_RET")
         .MoveNext
       Loop
      .MoveFirst
   End If
  End With
  'LabelTotal.Caption = Format(Total, "#,##0.00")
  LabelTotalRet.Caption = Format(Total_DetRet, "#,##0.00")
  DLTipoRet.SetFocus
End Sub

Private Sub TextValorR_GotFocus()
  TextValorR.Text = ""
  Label18.Caption = "Valor M/N"
  If Moneda_US Or OpcTM = 2 Then Label18.Caption = "Valor M/E"
End Sub

Private Sub TextValorR_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape: Command2.SetFocus
    Case vbKeyReturn: SiguienteControl
  End Select
End Sub

Private Sub TextValorR_LostFocus()
  TextoValido TextValorR, True
 'If Moneda_US Or OpcTM = 2 Then  TextValorR.Text = Round(Val(TextValorR.Text) * Dolar,2)
  TextValorR.Text = Round(TextValorR.Text, 2)
End Sub

Private Sub TxtCompRet_GotFocus()
  MarcarTexto TxtCompRet
End Sub

Private Sub TxtCompRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCompRet_LostFocus()
  TextoValido TxtCompRet, , True
End Sub
