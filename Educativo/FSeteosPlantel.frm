VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FSeteosPlantel 
   Caption         =   "DATOS GENERALES DE LA INSTITUCION"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSeteosPlantel.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSeteosPlantel.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSeteosPlantel.frx":11B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Frame1 
      Height          =   2010
      Left            =   105
      ScaleHeight     =   1950
      ScaleWidth      =   5100
      TabIndex        =   2
      Top             =   420
      Width           =   5160
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2535
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Tag             =   "Uno"
      Top             =   0
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   4471
      MultiRow        =   -1  'True
      TabFixedWidth   =   4
      TabFixedHeight  =   4
      Separators      =   -1  'True
      TabStyle        =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Uno"
            Key             =   "Uno"
            Object.Tag             =   "Uno"
            Object.ToolTipText     =   "Hola Uno"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dos"
            Key             =   "Dos"
            Object.Tag             =   "Dos"
            Object.ToolTipText     =   "Hola Dos"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tres"
            Key             =   "Tres"
            Object.Tag             =   "Tres"
            Object.ToolTipText     =   "hola Tres"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox CommandButton1 
      Height          =   540
      Left            =   5145
      ScaleHeight     =   480
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   3465
      Width           =   2115
   End
End
Attribute VB_Name = "FSeteosPlantel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''Item|NVARCHAR (3) NULL|TEXT (3) NULL|
'''Director|NVARCHAR (30) NULL|TEXT (30) NULL|
'''Secretario1|NVARCHAR (30) NULL|TEXT (30) NULL|
'''Rector|NVARCHAR (30) NULL|TEXT (30) NULL|
'''Secretario2|NVARCHAR (30) NULL|TEXT (30) NULL|
'''Anio_Lectivo|NVARCHAR (12) NULL|TEXT (12) NULL|
'''Periodo|NVARCHAR (10) NULL|TEXT (10) NULL|
'''Recomendacion|BIT NULL|BIT NULL|
'''Escala|BIT NULL|BIT NULL|
'''NPQP1|BIT NULL|BIT NULL|
'''NPQP2|BIT NULL|BIT NULL|
'''NPQEX|BIT NULL|BIT NULL|
'''NSQP1|BIT NULL|BIT NULL|
'''NSQP2|BIT NULL|BIT NULL|
'''NSQEX|BIT NULL|BIT NULL|
'''NSUPL|BIT NULL|BIT NULL|
'''NGRADO|BIT NULL|BIT NULL|
'''Formato|NVARCHAR (10) NULL|TEXT (10) NULL|
'''Texto_Director|NVARCHAR (15) NULL|TEXT (15) NULL|
'''Texto_Rector|NVARCHAR (15) NULL|TEXT (15) NULL|
'''Secretario3|NVARCHAR (30) NULL|TEXT (30) NULL|
'''Bachiller1|NVARCHAR (50) NULL|TEXT (50) NULL|
'''Bachiller2|NVARCHAR (50) NULL|TEXT (50) NULL|
'''Vicerrector1|NVARCHAR (30) NULL|TEXT (30) NULL|
'''Vicerrector2|NVARCHAR (30) NULL|TEXT (30) NULL|
'''Alfabetico|BIT NULL|BIT NULL|
'''Institucion1|NVARCHAR (50) NULL|TEXT (50) NULL|
'''Institucion2|NVARCHAR (50) NULL|TEXT (50) NULL|
'''Logo_Tipo|NVARCHAR (8) NULL|TEXT (8) NULL|
'''Texto_Secretario1|NVARCHAR (50) NULL|TEXT (50) NULL|
'''Texto_Secretario2|NVARCHAR (50) NULL|TEXT (50) NULL|
'''Rubro_Matricula|NVARCHAR (16) NULL|TEXT (16) NULL|
'''Codigo_Colegio|NVARCHAR (4) NULL|TEXT (4) NULL|
'''Codigo_AMIE|NVARCHAR (10) NULL|TEXT (10) NULL|
'''Mail_Colegio|NVARCHAR (50) NULL|TEXT (50) NULL|
'''Dec_Nota|TINYINT NULL|BYTE NULL|
'''Encabezado_Prim|BIT NULL|BIT NULL|
'''Encabezado_Secu|BIT NULL|BIT NULL|
'''Encabezado_Bach|BIT NULL|BIT NULL|

Private mintCurFrame As Integer ' Marco activo visible

Private Sub Form_Activate()
  RatonNormal
End Sub

Private Sub Tabstrip1_Click()
   If TabStrip1.SelectedItem.Index = mintCurFrame Then Exit Sub       ' No necesita cambiar el marco.
   ' Oculte el marco antiguo y muestre el nuevo.
   Frame1(TabStrip1.SelectedItem.Index).Visible = True
   Frame1(mintCurFrame).Visible = False
   ' Establece mintCurFrame al nuevo valor.
   mintCurFrame = TabStrip1.SelectedItem.Index
End Sub

