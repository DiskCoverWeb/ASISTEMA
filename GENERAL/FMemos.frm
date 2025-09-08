VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FMemos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MEMORANDO"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "FMemos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   1588
      ButtonWidth     =   2143
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Grabar"
            Key             =   "Grabar"
            Object.ToolTipText     =   "Graba el Memo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Imprimir"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir Memo"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Email"
            Key             =   "Email"
            Object.ToolTipText     =   "Enviar por mail"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Mails Todos"
            Key             =   "Todos_Mails"
            Object.ToolTipText     =   "Enviar el mail a todos los beneficiarios"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCCC1 
      Bindings        =   "FMemos.frx":0696
      DataSource      =   "AdoPara"
      Height          =   330
      Left            =   1050
      TabIndex        =   11
      Top             =   2205
      Visible         =   0   'False
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCNumero 
      Bindings        =   "FMemos.frx":06AC
      DataSource      =   "AdoMemoNo"
      Height          =   345
      Left            =   9345
      TabIndex        =   3
      Top             =   945
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox CheqMemo 
      BackColor       =   &H00FFC0C0&
      Caption         =   " NUMERO &No."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7665
      TabIndex        =   2
      Top             =   945
      Width           =   1695
   End
   Begin VB.TextBox TxtArchivo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2205
      TabIndex        =   21
      Top             =   3570
      Width           =   8835
   End
   Begin VB.TextBox TxtMemo 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2850
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   4305
      Width           =   10935
   End
   Begin MSDataListLib.DataCombo DCCC2 
      Bindings        =   "FMemos.frx":06C4
      DataSource      =   "AdoPara"
      Height          =   330
      Left            =   1050
      TabIndex        =   14
      Top             =   2625
      Visible         =   0   'False
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox CheqCC2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CC"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   13
      Top             =   2625
      Width           =   960
   End
   Begin VB.CheckBox CheqCC1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CC"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   2205
      Width           =   960
   End
   Begin MSAdodcLib.Adodc AdoPara 
      Height          =   330
      Left            =   945
      Top             =   5880
      Visible         =   0   'False
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
      Caption         =   "Para"
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
   Begin VB.TextBox TxtAtencion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6825
      MaxLength       =   30
      TabIndex        =   19
      Top             =   3150
      Width           =   4215
   End
   Begin MSDataListLib.DataCombo DCPara 
      Bindings        =   "FMemos.frx":06DA
      DataSource      =   "AdoPara"
      Height          =   330
      Left            =   1050
      TabIndex        =   8
      Top             =   1785
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtAsunto 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1155
      MaxLength       =   30
      TabIndex        =   17
      Top             =   3150
      Width           =   4320
   End
   Begin MSAdodcLib.Adodc AdoMemoNo 
      Height          =   330
      Left            =   945
      Top             =   6195
      Visible         =   0   'False
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
      Caption         =   "MemoNo"
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
   Begin MSAdodcLib.Adodc AdoMemo 
      Height          =   330
      Left            =   945
      Top             =   6510
      Visible         =   0   'False
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
      Caption         =   "Memo"
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
   Begin MSComDlg.CommonDialog CDialogDir 
      Left            =   9975
      Top             =   7245
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Archivo Adjunto:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   20
      Top             =   3570
      Width           =   2115
   End
   Begin VB.Label LblEmailCC2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   5565
      TabIndex        =   15
      Top             =   2625
      Visible         =   0   'False
      Width           =   5475
   End
   Begin VB.Label LblEmailCC1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   5565
      TabIndex        =   12
      Top             =   2205
      Visible         =   0   'False
      Width           =   5475
   End
   Begin VB.Label LblEmailPara 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   5565
      TabIndex        =   9
      Top             =   1785
      Width           =   5475
   End
   Begin VB.Label LblEmailDe 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   5565
      TabIndex        =   6
      Top             =   1365
      Width           =   5475
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   10500
      Top             =   7245
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FMemos.frx":06F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FMemos.frx":0A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FMemos.frx":0D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FMemos.frx":0F62
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FMemos.frx":127C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ATENCION:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   5565
      TabIndex        =   18
      Top             =   3150
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PARA:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ASUNTO:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   16
      Top             =   3150
      Width           =   1065
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1050
      TabIndex        =   5
      Top             =   1365
      Width           =   4530
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ATENCION:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   1050
      TabIndex        =   1
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   945
      Width           =   960
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " M E M O R A N D O"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   22
      Top             =   3990
      Width           =   10935
   End
   Begin VB.Label LblDe 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   105
      TabIndex        =   25
      Top             =   7560
      Width           =   6000
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Atentamente,"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   105
      TabIndex        =   24
      Top             =   7245
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DE:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   1365
      Width           =   960
   End
End
Attribute VB_Name = "FMemos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheqCC1_Click()
  If CheqCC1.value = 1 Then
     DCCC1.Visible = True
     LblEmailCC1.Visible = True
  Else
     DCCC1.Visible = False
     LblEmailCC1.Visible = False
  End If
End Sub

Private Sub CheqCC2_Click()
  If CheqCC2.value = 1 Then
     DCCC2.Visible = True
     LblEmailCC2.Visible = True
  Else
     DCCC2.Visible = False
     LblEmailCC2.Visible = False
  End If
End Sub

Public Sub Grabar_Memo()
  TextoValido TxtAsunto, , True
  TextoValido TxtAtencion, , True
  If TxtMemo = "" Then TxtMemo = Ninguno
  If CheqCC1.value <> 1 Then Codigo1 = Ninguno
  If CheqCC2.value <> 1 Then Codigo2 = Ninguno
  If Len(TxtAsunto) > 1 And Len(TxtAtencion) > 1 And Len(TxtMemo) > 1 Then
     Mensajes = "Esta seguro de Grabar Memo"
     Titulo = "Pregunta de grabación"
     If BoxMensaje = vbYes Then
        Numero = ReadSetDataNum("Memos", True, True)
        sSQL = "DELETE * " _
             & "FROM Trans_Memos " _
             & "WHERE Numero = " & Numero & " " _
             & "AND Item = '" & NumEmpresa & "' "
        Ejecutar_SQL_SP sSQL
        
        SetAdoAddNew "Trans_Memos"
        SetAdoFields "Fecha", FechaSistema
        SetAdoFields "Hora", Format(Time, FormatoTimes)
        SetAdoFields "Texto_Memo", TxtMemo & vbCrLf & "   "
        SetAdoFields "Numero", Numero
        SetAdoFields "Asunto", TxtAsunto
        SetAdoFields "Atencion", TxtAtencion
        SetAdoFields "Codigo", CodigoCliente
        SetAdoFields "CC1", Codigo1
        SetAdoFields "CC2", Codigo2
        SetAdoFields "CodigoU", CodigoUsuario
        SetAdoFields "Periodo", Periodo_Contable
        SetAdoFields "Item", NumEmpresa
        SetAdoUpdate
        MsgBox "Grabado Completo"
        Imprimir_Memo AdoPara, AdoMemo, Numero
        Numero_Memos
     End If
  Else
     MsgBox "No se puede grabar, Faltan datos"
  End If
End Sub

Private Sub CheqMemo_Click()
  If CheqMemo = 1 Then DCNumero.Visible = True Else DCNumero.Visible = False
End Sub

Private Sub DCCC1_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCC1_LostFocus()
  Codigo1 = Ninguno
  LblEmailCC1.Caption = ""
  With AdoPara.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCCC1 & "' ")
       If Not .EOF Then
          Codigo1 = .Fields("Codigo")
          LblEmailCC1.Caption = .Fields("Email")
       End If
   End If
  End With
End Sub

Private Sub DCCC2_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub DCCC2_LostFocus()
  Codigo2 = Ninguno
  LblEmailCC2.Caption = ""
  With AdoPara.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCCC2 & "' ")
       If Not .EOF Then
          Codigo2 = .Fields("Codigo")
          LblEmailCC2.Caption = .Fields("Email")
       End If
   End If
  End With
End Sub

Private Sub DCNumero_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCNumero_LostFocus()
  Numero = Val(DCNumero)
  sSQL = "SELECT A.Nombre_Completo,C.Cliente,C.Email,TM.* " _
       & "FROM Accesos As A,Clientes As C,Trans_Memos As TM " _
       & "WHERE TM.Numero = " & Numero & " " _
       & "AND TM.Item = '" & NumEmpresa & "' " _
       & "AND A.Codigo = TM.CodigoU " _
       & "AND C.Codigo = TM.Codigo "
  Select_Adodc AdoMemo, sSQL
  With AdoMemo.Recordset
   If .RecordCount > 0 Then
       TxtAsunto = .Fields("Asunto")
       TxtAtencion = .Fields("Atencion")
       TxtMemo = .Fields("Texto_Memo")
       Label8.Caption = .Fields("Fecha")
       DCPara = .Fields("Cliente")
       LblEmailPara.Caption = .Fields("Email")
       LblDe.Caption = .Fields("Nombre_Completo")
       Label9.Caption = .Fields("Nombre_Completo")
       Codigo1 = .Fields("CC1")
       Codigo2 = .Fields("CC2")
       DCCC1.Visible = False: CheqCC1.value = 0
       DCCC2.Visible = False: CheqCC2.value = 0
       If AdoPara.Recordset.RecordCount > 0 Then
          AdoPara.Recordset.MoveFirst
          AdoPara.Recordset.Find ("Codigo = '" & Codigo1 & "' ")
          If Not AdoPara.Recordset.EOF Then
             DCCC1 = AdoPara.Recordset.Fields("Cliente")
             LblEmailCC1.Caption = AdoPara.Recordset.Fields("Email")
             DCCC1.Visible = True
             CheqCC1.value = 1
          End If
       End If
       If AdoPara.Recordset.RecordCount > 0 Then
          AdoPara.Recordset.MoveFirst
          AdoPara.Recordset.Find ("Codigo = '" & Codigo2 & "' ")
          If Not AdoPara.Recordset.EOF Then
             DCCC2 = AdoPara.Recordset.Fields("Cliente")
             LblEmailCC2.Caption = AdoPara.Recordset.Fields("Email")
             DCCC2.Visible = True
             CheqCC2.value = 1
          End If
       End If
       TxtMemo.SetFocus
   End If
  End With
End Sub

Private Sub DCPara_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCPara_LostFocus()
  CodigoCliente = Ninguno
  Codigo1 = Ninguno
  Codigo2 = Ninguno
  LblEmailPara.Caption = ""
  With AdoPara.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & DCPara & "' ")
       If Not .EOF Then
          CodigoCliente = .Fields("Codigo")
          LblEmailPara.Caption = .Fields("Email")
       End If
   End If
  End With
End Sub

Private Sub Form_Activate()
Dim X As Boolean
  Numero = ReadSetDataNum("Memos", True, False)
  Numero_Memos
  sSQL = "SELECT Codigo,Cliente,Email,Email2,Direccion,Telefono " _
       & "FROM Clientes " _
       & "WHERE Codigo <> '.' " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCPara, AdoPara, sSQL, "Cliente"
  SelectDB_Combo DCCC1, AdoPara, sSQL, "Cliente"
  SelectDB_Combo DCCC2, AdoPara, sSQL, "Cliente"
  LblDe.Caption = NombreUsuario
  Label8.Caption = FechaSistema
  Label9.Caption = NombreUsuario
  LblEmailDe.Caption = ""
  With AdoPara.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Cliente = '" & UCaseStrg(NombreUsuario) & "' ")
       If Not .EOF Then LblEmailDe.Caption = .Fields("Email")
   End If
  End With
  DCPara.SetFocus
  RatonNormal
End Sub

Private Sub Form_Load()
  CentrarForm FMemos
  ConectarAdodc AdoPara
  ConectarAdodc AdoMemo
  ConectarAdodc AdoMemoNo
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
 Select Case Button.key
   Case "Grabar"
        Grabar_Memo
   Case "Imprimir"
        Numero = Val(DCNumero)
        Imprimir_Memo AdoPara, AdoMemo, Numero
   Case "Email"
        TMail.Adjunto = ""
        If Len(TxtArchivo) > 1 Then TMail.Adjunto = TxtArchivo
        Email_Memo AdoPara, AdoMemo, Numero
   Case "Todos_Mails"
        Enviar_mails_Todos
   Case "Salir"
        Unload FMemos
 End Select
End Sub

Private Sub TxtArchivo_DblClick()
  TxtArchivo = SelectDialogFile(CDialogDir)
  TxtMemo.SetFocus
End Sub

Private Sub TxtArchivo_GotFocus()
  MarcarTexto TxtArchivo
End Sub

Private Sub TxtArchivo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAsunto_GotFocus()
  MarcarTexto TxtAsunto
End Sub

Private Sub TxtAsunto_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAsunto_LostFocus()
  TextoValido TxtAsunto, , True
End Sub

Private Sub TxtAtencion_GotFocus()
  MarcarTexto TxtAtencion
End Sub

Private Sub TxtAtencion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtAtencion_LostFocus()
  TextoValido TxtAtencion, , True
End Sub

Private Sub TxtMemo_GotFocus()
  MarcarTexto TxtMemo
End Sub

Public Sub Numero_Memos()
  sSQL = "SELECT * " _
       & "FROM Trans_Memos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "ORDER BY Numero "
  SelectDB_Combo DCNumero, AdoMemoNo, sSQL, "Numero"
End Sub

Public Sub Enviar_mails_Todos()
Dim LosMails As String
Dim PuntoyComa As Integer
    TMail.ListaMail = 0
    TMail.Adjunto = ""
    If Len(TxtArchivo) > 1 Then TMail.Adjunto = TxtArchivo
    TMail.Asunto = TxtAsunto
    TMail.Mensaje = TxtMemo
    TMail.Mensaje = TMail.Mensaje _
                 & " " & vbCrLf _
                 & " " & vbCrLf _
                 & String(80, "_") & vbCrLf _
                 & "Este mensaje enviado a Clientes de DiskCover System, " & vbCrLf _
                 & "si no desea recibir mas correos envie un mail con la " & vbCrLf _
                 & "palabra NO RECIBIR MAS " & vbCrLf
    With AdoPara.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
         Do While Not .EOF
            LosMails = ""
            If Len(.Fields("Email")) > 1 Then LosMails = LosMails & .Fields("Email") & ";"
            If Len(.Fields("Email2")) > 1 Then LosMails = LosMails & .Fields("Email2") & ";"
            LosMails = TrimStrg(LosMails)
            Do While Len(LosMails) > 2
               PuntoyComa = InStr(LosMails, ";")
               TMail.para = MidStrg(LosMails, 1, PuntoyComa - 1)
               If EsUnEmail(TMail.para) Then FEnviarCorreos.Show 1
               MsgBox MidStrg(LosMails, PuntoyComa + 1, Len(LosMails))
               LosMails = MidStrg(LosMails, PuntoyComa + 1, Len(LosMails))
            Loop
           .MoveNext
         Loop
         MsgBox "Proceso Exitoso"
     End If
    End With
End Sub
