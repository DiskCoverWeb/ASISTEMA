VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form IngFechas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1155
   ClientLeft      =   30
   ClientTop       =   -15
   ClientWidth     =   3975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Fechas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3795
      Begin VB.CommandButton Command2 
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
         Height          =   645
         Left            =   1470
         Picture         =   "Fechas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
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
         Height          =   645
         Left            =   2625
         Picture         =   "Fechas.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1065
      End
      Begin MSMask.MaskEdBox MBFecha 
         Height          =   330
         Left            =   105
         TabIndex        =   2
         Top             =   525
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   1275
      End
   End
   Begin MSAdodcLib.Adodc AdoSup 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Sup"
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
Attribute VB_Name = "IngFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
  Unload IngFechas
End Sub

Private Sub Command2_Click()

 If IsDate(MBFecha) Then

   'Actualizar las Ctas a mayoriazar
    sSQL = "SELECT Cta " _
         & "FROM Transacciones " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND TP = '" & Co.TP & "' " _
         & "AND Numero = " & Co.Numero & " " _
         & "GROUP BY Cta " _
         & "ORDER BY Cta "
    Select_Adodc AdoSup, sSQL
    With AdoSup.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
           'Determinamos que la cuenta ya fue mayorizada
            SubCta = .fields("Cta")
            sSQL = "UPDATE Transacciones " _
                 & "SET Procesado = " & Val(adFalse) & " " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Cta = '" & SubCta & "' "
            Ejecutar_SQL_SP sSQL
           .MoveNext
         Loop
     End If
    End With
    sSQL = "SELECT Codigo_Inv " _
         & "FROM Trans_Kardex " _
         & "WHERE Periodo = '" & Periodo_Contable & "' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND TP = '" & Co.TP & "' " _
         & "AND Numero = " & Co.Numero & " " _
         & "GROUP BY Codigo_Inv " _
         & "ORDER BY Codigo_Inv "
    Select_Adodc AdoSup, sSQL
    With AdoSup.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
           'Determinamos que la cuenta ya fue mayorizada
            SubCta = .fields("Codigo_Inv")
            sSQL = "UPDATE Trans_Kardex " _
                 & "SET Procesado = " & Val(adFalse) & " " _
                 & "WHERE Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "AND Codigo_Inv = '" & SubCta & "' "
            Ejecutar_SQL_SP sSQL
           .MoveNext
         Loop
     End If
    End With
    MsgBox "Proceso Terminado, Vuelva a Mayorizar"
 Else
    MsgBox "Fecha invalida, no se procedera hacer nada"
 End If
 Unload IngFechas
End Sub

Private Sub Form_Load()
  CentrarForm IngFechas
  IngFechas.Caption = "TIPO DE COMPROBANTE: " & Co.TP & " No. " & Format(Co.Numero, "00000000")
  ConectarAdodc AdoSup
  RatonNormal
End Sub

Private Sub MBFecha_GotFocus()
   MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub MBFecha_LostFocus()
   FechaValida MBFecha
End Sub
