VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form SetPrinters 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   Caption         =   "Seleccione la Impresora"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "SetPrinters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoSetImpresora 
      Height          =   330
      Left            =   6405
      Top             =   3255
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
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
      Caption         =   "SetImpresora"
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
   Begin VB.OptionButton OpcBHoriz 
      BackColor       =   &H00FFC0C0&
      Height          =   645
      Left            =   6615
      Picture         =   "SetPrinters.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   945
      Width           =   960
   End
   Begin VB.OptionButton OpcBVert 
      BackColor       =   &H00FFC0C0&
      Height          =   750
      Left            =   6720
      Picture         =   "SetPrinters.frx":100C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   105
      Value           =   -1  'True
      Width           =   750
   End
   Begin VB.TextBox TxtPapel 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   105
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2205
      Width           =   5055
   End
   Begin VB.TextBox TxtImpresora 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   105
      TabIndex        =   3
      Top             =   1890
      Width           =   6210
   End
   Begin VB.ListBox ListPRN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   105
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   6210
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
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
      Height          =   750
      Left            =   6405
      Picture         =   "SetPrinters.frx":174E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
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
      Height          =   750
      Left            =   6405
      Picture         =   "SetPrinters.frx":1DE4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1275
   End
   Begin VB.ListBox ListPrinter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   6210
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Detalles>>"
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
      TabIndex        =   5
      Top             =   2205
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECCIONE LA IMPRESORA A IMPRIMIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   6210
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SELECCIONE LA IMPRESORA"
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
      TabIndex        =   0
      Top             =   105
      Width           =   6210
   End
End
Attribute VB_Name = "SetPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   SetNombrePRN = ListPrinter.Text
   If SetNombrePRN = Impresota_PDF Then TxtPapel = "009 - A4 (21.0 x 29.7)" Else TxtPapel = ListPRN
   SetPapelPRNCad = TxtPapel
   SetPapelPRN = CInt(SinEspaciosIzq(TxtPapel))
  'MsgBox TxtPapel
   If OpcBVert.value Then Orientacion_Pagina = 1 Else Orientacion_Pagina = 2
   SetPapelCopia = False
   With AdoSetImpresora.Recordset
    If .RecordCount > 0 Then
       .Fields("Impresora_Defecto") = TxtImpresora
       .Fields("Papel_Impresora") = TxtPapel
       .Update
    End If
   End With
   SetPrinters.Hide
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Presiono_Esc
End Sub

Private Sub Command2_Click()
  Presiono_Esc
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Presiono_Esc
End Sub

Private Sub Command3_Click()
  If ListPRN.Visible Then
     ListPRN.Visible = False
     ExpandirPRN = False
     Command3.Caption = "&Detalle >>"
  Else
     ListPRN.Visible = True
     ExpandirPRN = True
     Command3.Caption = "<< &Detalle"
  End If
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Presiono_Esc
End Sub

Private Sub Form_Activate()
  If CodigoUsuario = "" Then CodigoUsuario = Ninguno
 'Abrimos el usuario por dedault
  sSQL = "SELECT Codigo,Impresora_Defecto,Papel_Impresora,ID " _
       & "FROM Accesos " _
       & "WHERE Codigo = '" & CodigoUsuario & "' "
  Select_Adodc AdoSetImpresora, sSQL
 
  If ListPrinter.ListCount > 0 Then
     TxtImpresora = ListPrinter.Text
     SetNombrePRN = ListPrinter.Text
     Si_No = PapelesImpresora(ListPrinter.Text, ListPRN)
    'MsgBox ListPRN.ListCount
     If ListPRN.ListCount > 0 Then
        SetPapelPRNCad = ListPRN.List(0)
        For i = 0 To ListPRN.ListCount - 1
            If Val(SinEspaciosIzq(ListPRN.List(i))) = 9 Then SetPapelPRNCad = ListPRN.List(i)
        Next i
     Else
        SetPapelPRNCad = "No Admite tipo y porte de papel"
     End If
     If ListPRN.Text = "" Then ListPRN.Text = SetPapelPRNCad
     TxtPapel = SetPapelPRNCad
     If Orientacion_Pagina < 1 Then Orientacion_Pagina = 1
     If Orientacion_Pagina = 1 Then OpcBVert.value = True Else OpcBHoriz.value = True
     If Titulo <> "" Then SetPrinters.Caption = Titulo Else SetPrinters.Caption = "IMPRIMIR"
     If Mensajes <> "" Then Label2.Caption = Mensajes
    'Determinamos si esta organizado por default el tipo de impresora y tamaño de papel
     If AdoSetImpresora.Recordset.RecordCount > 0 Then
        If Len(AdoSetImpresora.Recordset.Fields("Impresora_Defecto")) > 1 Then
           ListPrinter.Text = AdoSetImpresora.Recordset.Fields("Impresora_Defecto")
           TxtPapel = AdoSetImpresora.Recordset.Fields("Papel_Impresora")
        End If
     End If
     SetPrinters.Visible = True
     ListPRN.Visible = True
     Command1.SetFocus
     RatonNormal
     If ListPRN.Visible Then
        ListPRN.Visible = False
        ExpandirPRN = False
        Command3.Caption = "&Detalle >>"
     Else
        ListPRN.Visible = True
        ExpandirPRN = True
        Command3.Caption = "<< &Detalle"
     End If
     'MsgBox ListPRN.Text
  Else
     RatonNormal
     MsgBox "No tiene Impresora conectada," & vbCrLf _
          & "Instale por lo menos una para" & vbCrLf _
          & "poder utilizar el Sistema"
     End
  End If
End Sub

Private Sub Form_Load()
  RatonReloj
  CentrarForm SetPrinters
  ConectarAdodc AdoSetImpresora
 
 'Revizamos cuantas Impresoras estan instaladas y activas
  ReDim ListaDeImpresoras(Printers.Count + 1) As String
  For i = 0 To Printers.Count - 1
      ListaDeImpresoras(i) = Printers(i).DeviceName
  Next i
  ListaDeImpresoras(Printers.Count) = Impresota_PDF
  
 'llenamos las impresoras en el combobox
  ListPrinter.Clear
  For i = 0 To UBound(ListaDeImpresoras) - 1
      ListPrinter.AddItem ListaDeImpresoras(i)
  Next i
  ListPrinter.Text = Printer.DeviceName
End Sub

Private Sub ListPrinter_DblClick()
  SiguienteControl
End Sub

Private Sub ListPrinter_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Presiono_Esc
End Sub

Private Sub ListPrinter_LostFocus()
   
  Si_No = PapelesImpresora(ListPrinter.Text, ListPRN)
    'MsgBox ListPRN.ListCount
     If ListPRN.ListCount > 0 Then
        SetPapelPRNCad = ListPRN.List(0)
        For i = 0 To ListPRN.ListCount - 1
            If Val(SinEspaciosIzq(ListPRN.List(i))) = 9 Then SetPapelPRNCad = ListPRN.List(i)
        Next i
     Else
        SetPapelPRNCad = "No Admite tipo y porte de papele"
     End If
     If ListPRN.Text = "" Then ListPRN.Text = SetPapelPRNCad
     TxtImpresora = ListPrinter.Text
     TxtPapel = SetPapelPRNCad
  'MsgBox SetPapelPRNCad
'  LblPapel.Caption = ListPRN.Text
End Sub

Private Sub ListPRN_Click()
  TxtPapel = ListPRN.Text
End Sub

Private Sub ListPRN_DblClick()
  TxtPapel = ListPRN.Text
  SiguienteControl
End Sub

Public Sub Presiono_Esc()
   TxtPapel = ListPRN.Text
   If OpcBVert.value Then Orientacion_Pagina = 1 Else Orientacion_Pagina = 2
   SetPapelPRNCad = ListPRN.Text
   SetPapelPRN = CInt(SinEspaciosIzq(ListPRN.Text))
   SetNombrePRN = ""
   SetPrinters.Hide
End Sub

Private Sub ListPRN_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
  If KeyCode = vbKeyEscape Then Presiono_Esc
End Sub

