VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form FVerPDF 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   7800
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar TBarCliente 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Módulo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Enviar a Excel la consulta"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Emails"
            Object.ToolTipText     =   "Enviar por mail el PDF"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame FrmTotal 
         Height          =   645
         Left            =   1890
         TabIndex        =   2
         Top             =   0
         Width           =   3585
         Begin VB.TextBox TxtHSaldoActual 
            Alignment       =   1  'Right Justify
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
            Left            =   1785
            MaxLength       =   14
            TabIndex        =   3
            Text            =   "0.00"
            Top             =   210
            Width           =   1695
         End
         Begin VB.Label Label32 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " SALDO ACTUAL"
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
            TabIndex        =   4
            Top             =   210
            Width           =   1695
         End
      End
   End
   Begin AcroPDFLibCtl.AcroPDF ArchivoPDF 
      Height          =   7470
      Left            =   105
      TabIndex        =   0
      Top             =   735
      Width           =   12195
      _cx             =   5080
      _cy             =   5080
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   16380
      Top             =   1155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FVerPDF.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FVerPDF.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FVerPDF.frx":0F6C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FVerPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim Resultado As Boolean

    If VerPDF.CodigoBeneficiario = "" Then VerPDF.CodigoBeneficiario = Ninguno
    VerPDF.EmailBeneficiario = ""
    Insertar_Mail VerPDF.EmailBeneficiario, TBeneficiario.EmailR
    Insertar_Mail VerPDF.EmailBeneficiario, TBeneficiario.Email2
    Insertar_Mail VerPDF.EmailBeneficiario, TBeneficiario.Email1
   ' MsgBox VerPDF.CodigoBeneficiario
    Select Case VerPDF.TipoPDF
      Case "HISTORIAL"
           Resultado = Reporte_Cartera_Clientes_PDF(FechaInicial, VerPDF.CodigoBeneficiario, ArchivoPDF, False)
           If Resultado Then
              FVerPDF.Caption = VerPDF.Titulo & "-" & VerPDF.TipoPDF & "-" & VerPDF.CodigoBeneficiario & " -> " & RutaDocumentoPDF
              MsgBox "Este archivo fue generado es:" & vbCrLf & vbCrLf & RutaDocumentoPDF
              If Len(RutaDocumentoPDF) > 1 Then
                 Presentar_PDF ArchivoPDF, RutaDocumentoPDF, 125
                 If VerPDF.ValorTotal > 0 Then
                    FrmTotal.Visible = True
                    TxtHSaldoActual = Format(VerPDF.ValorTotal, "#,##0.00")
                 Else
                    FrmTotal.Visible = False
                 End If
              End If
              RatonNormal
           Else
              RatonNormal
              MsgBox "No hay datos que Procesar"
              Unload FVerPDF
           End If
      Case Else
           RatonNormal
    End Select
End Sub

Private Sub Form_Load()

    'FVerPDF.Caption = VerPDF.Titulo & "-" & VerPDF.TipoPDF & "-" & VerPDF.CodigoBeneficiario
    ArchivoPDF.width = MDI_X_Max - 50
    ArchivoPDF.Height = MDI_Y_Max - 300
End Sub

Private Sub TBarCliente_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
      Case "Salir"
           Unload FVerPDF
      Case "Excel"
           Select Case VerPDF.TipoPDF
             Case "HISTORIAL"
                  Reporte_Cartera_Clientes_PDF FechaInicial, VerPDF.CodigoBeneficiario, ArchivoPDF, True
           End Select
      Case "Emails"
''    ComunicadoEntidad = ""
''    CodigoP = Format$(Total, "#,#0.00")
''    CodigoP = String(14 - Len(CodigoP), " ") & CodigoP
      
           TMail.TipoDeEnvio = Ninguno
           TMail.Asunto = "Estimado(a): " & VerPDF.NombreBeneficiario & ", usted tiene los siguientes pendientes."
           TMail.Mensaje = "Envio automatizado de su cartera pendiente." & vbCrLf _
                         & "NOTA: En caso de tener inconformidad con los valores detallados en su Estado de Cuenta, comuniquese con atencion al Cliente."
           TMail.Adjunto = RutaDocumentoPDF
           TMail.para = VerPDF.EmailBeneficiario
          'Enviamos lista de mails
           FEnviarCorreos.Show vbModal
    End Select
End Sub
