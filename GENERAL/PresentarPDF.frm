VERSION 5.00
Begin VB.Form FPresentarPDF 
   BackColor       =   &H00800080&
   Caption         =   "Cargar PDF en Formulario VB"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin VB.PictureBox WBPDF 
      Height          =   9885
      Left            =   210
      ScaleHeight     =   9825
      ScaleWidth      =   14865
      TabIndex        =   2
      Top             =   630
      Width           =   14925
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
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
      Height          =   390
      Left            =   11235
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UTILICE LOS CONTROLES PROPIOS DEL NAVEGADOR PARA IMPRIMIR, GUARDAR  O ENVIAR POR MAIL EL ARCHIVO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   210
      Width           =   10725
   End
End
Attribute VB_Name = "FPresentarPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Unload FPresentarPDF
End Sub

Private Sub Form_Load()
   RatonReloj
   WBPDF.Height = MDI_Y_Max - WBPDF.Top - 100
   WBPDF.width = MDI_X_Max - 250
   Command1.Left = WBPDF.width - 900

   WBPDF.Navigate "file:///C:/SISTEMA/FONDOS/index_pdf.html"

   If Existe_File(RutaDocumentoPDF) Then
      WBPDF.Navigate RutaDocumentoPDF
      RatonNormal
   Else
      RatonNormal
      MsgBox "No existe archivo que presentar, revise que de verdad existe."
      Unload FPresentarPDF
   End If
End Sub

