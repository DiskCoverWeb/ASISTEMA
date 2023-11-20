VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FCatalogo_Cuentas1 
   Caption         =   "Ingreso de Cuentas Contables"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmPresupuesto 
      BackColor       =   &H00C0FFFF&
      Caption         =   ":<| PRESUPUESTO |>:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   18795
      TabIndex        =   36
      Top             =   1890
      Visible         =   0   'False
      Width           =   3060
      Begin VB.TextBox TxtPresMes 
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
         Index           =   0
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   50
         Text            =   "0.00"
         Top             =   420
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Actualizar"
         Height          =   330
         Left            =   840
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   49
         Top             =   4305
         Width           =   960
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Cancelar"
         Height          =   330
         Left            =   1995
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   48
         Top             =   4305
         Width           =   960
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   1
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   47
         Text            =   "0.00"
         Top             =   735
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   2
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   46
         Text            =   "0.00"
         Top             =   1050
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   3
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   1365
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   4
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   44
         Text            =   "0.00"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   5
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   43
         Text            =   "0.00"
         Top             =   1995
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   6
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   2310
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   7
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   2625
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   8
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   40
         Text            =   "0.00"
         Top             =   2940
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   9
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   3255
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   10
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   3570
         Width           =   1695
      End
      Begin VB.TextBox TxtPresMes 
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
         Index           =   11
         Left            =   1260
         MaxLength       =   12
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   3885
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   0
         Left            =   105
         TabIndex        =   62
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   1
         Left            =   105
         TabIndex        =   61
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   2
         Left            =   105
         TabIndex        =   60
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   3
         Left            =   105
         TabIndex        =   59
         Top             =   1365
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   4
         Left            =   105
         TabIndex        =   58
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   5
         Left            =   105
         TabIndex        =   57
         Top             =   1995
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   6
         Left            =   105
         TabIndex        =   56
         Top             =   2310
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   7
         Left            =   105
         TabIndex        =   55
         Top             =   2625
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   8
         Left            =   105
         TabIndex        =   54
         Top             =   2940
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   9
         Left            =   105
         TabIndex        =   53
         Top             =   3255
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   10
         Left            =   105
         TabIndex        =   52
         Top             =   3570
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Diciembre"
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
         Index           =   11
         Left            =   105
         TabIndex        =   51
         Top             =   3885
         Width           =   1065
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del modulo"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   18795
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":0BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":14CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":17E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":20C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":289C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":3176
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":4F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":5852
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":612C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":6446
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":6760
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":703A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":7914
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":7C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":8508
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":8822
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":8B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":8E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":92A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":9B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":9FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIngCtas1.frx":A2EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   7680
      Left            =   7875
      TabIndex        =   2
      Top             =   735
      Width           =   9780
      Begin VB.CheckBox CheqTipoPago 
         Caption         =   "TIPO DE PAGO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   32
         Top             =   4410
         Width           =   1485
      End
      Begin VB.CheckBox CheqFE 
         Caption         =   "Flujo de Efectivo"
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
         TabIndex        =   31
         Top             =   3885
         Width           =   1800
      End
      Begin VB.CheckBox CheqUS 
         Caption         =   "Cuenta M/E"
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
         TabIndex        =   30
         Top             =   3360
         Width           =   1275
      End
      Begin VB.CheckBox CheqModGastos 
         Caption         =   "Para Gastos de Caja Chica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   29
         Top             =   2730
         Width           =   1695
      End
      Begin VB.ListBox LstSubMod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2400
         Left            =   1995
         TabIndex        =   27
         Top             =   1995
         Width           =   3480
      End
      Begin VB.Frame Frame1 
         Caption         =   "Para Rol de Pagos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   5565
         TabIndex        =   22
         Top             =   1680
         Width           =   2010
         Begin VB.OptionButton OpcNoAplica 
            Caption         =   "No Aplica"
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
            TabIndex        =   26
            Top             =   315
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.CheckBox CheqConIESS 
            Caption         =   "Ingreso extra con Aplicacion al IESS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   105
            TabIndex        =   25
            Top             =   1260
            Width           =   1590
         End
         Begin VB.OptionButton OpcIEmp 
            Caption         =   "Ingreso"
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
            TabIndex        =   24
            Top             =   630
            Width           =   960
         End
         Begin VB.OptionButton OpcEEmp 
            Caption         =   "Descuentos"
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
            TabIndex        =   23
            Top             =   945
            Width           =   1380
         End
      End
      Begin VB.TextBox TextPresupuesto 
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
         Left            =   7665
         MaxLength       =   12
         TabIndex        =   18
         Text            =   "0.00"
         ToolTipText     =   "Ctrl+Insert: Insertar Presupuesto"
         Top             =   1995
         Width           =   2010
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   105
         TabIndex        =   15
         Top             =   1680
         Width           =   1800
         Begin VB.OptionButton OpcD 
            Caption         =   "Detalle"
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
            TabIndex        =   17
            Top             =   210
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OpcG 
            Caption         =   "Grupo"
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
            TabIndex        =   16
            Top             =   525
            Width           =   855
         End
      End
      Begin VB.TextBox TxtCodExt 
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
         Left            =   3780
         MaxLength       =   14
         TabIndex        =   13
         Text            =   "0"
         Top             =   1260
         Width           =   1695
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
         Left            =   1995
         MaxLength       =   90
         TabIndex        =   5
         Top             =   525
         Width           =   7680
      End
      Begin MSMask.MaskEdBox MBoxCta 
         Height          =   330
         Left            =   105
         TabIndex        =   6
         Top             =   525
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBoxCtaAcreditar 
         Height          =   330
         Left            =   6825
         TabIndex        =   20
         Top             =   1260
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo DCTipoPago 
         Bindings        =   "FIngCtas1.frx":ABC8
         DataSource      =   "AdoTipoPago"
         Height          =   315
         Left            =   1995
         TabIndex        =   33
         Top             =   4515
         Visible         =   0   'False
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648447
         Text            =   ""
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
      Begin MSDataGridLib.DataGrid DGGastos 
         Bindings        =   "FIngCtas1.frx":ABE2
         Height          =   2640
         Left            =   105
         TabIndex        =   34
         Top             =   4935
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   4657
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo de Cuenta"
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
         Left            =   1995
         TabIndex        =   28
         Top             =   1680
         Width           =   3480
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Acreditar"
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
         TabIndex        =   21
         Top             =   945
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Presupuesto"
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
         Left            =   7665
         TabIndex        =   19
         Top             =   1680
         Width           =   2010
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo Externo"
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
         Left            =   3780
         TabIndex        =   14
         Top             =   945
         Width           =   1695
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Numero"
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
         Left            =   5565
         TabIndex        =   12
         Top             =   945
         Width           =   1170
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
         Left            =   5565
         TabIndex        =   11
         Top             =   1260
         Width           =   1170
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
         Left            =   105
         TabIndex        =   10
         Top             =   945
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
         Left            =   105
         TabIndex        =   9
         Top             =   1260
         Width           =   1800
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo de Cuenta"
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
         Left            =   1995
         TabIndex        =   8
         Top             =   945
         Width           =   1695
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
         Left            =   1995
         TabIndex        =   7
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE DE LA CUENTA"
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
         Left            =   1995
         TabIndex        =   4
         Top             =   210
         Width           =   7680
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo de Cuenta"
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
         TabIndex        =   3
         Top             =   210
         Width           =   1800
      End
   End
   Begin MSComctlLib.TreeView TVCatalogo 
      Height          =   7260
      Left            =   105
      TabIndex        =   1
      Top             =   1155
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   12806
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc AdoCta 
      Height          =   330
      Left            =   735
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
   Begin MSAdodcLib.Adodc AdoGastos 
      Height          =   330
      Left            =   735
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
      Caption         =   "Gastos"
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
      Left            =   735
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
      Left            =   735
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
      Left            =   735
      Top             =   1470
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
   Begin MSAdodcLib.Adodc AdoPresupuesto 
      Height          =   330
      Left            =   735
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
      Caption         =   "Presupuesto"
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
   Begin MSAdodcLib.Adodc AdoGastos1 
      Height          =   330
      Left            =   735
      Top             =   3360
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
      Caption         =   "Gastos1"
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
   Begin MSAdodcLib.Adodc AdoTipoPago 
      Height          =   330
      Left            =   735
      Top             =   3675
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
      Caption         =   "TipoPago"
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
      Top             =   840
      Width           =   7680
   End
End
Attribute VB_Name = "FCatalogo_Cuentas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Cta_Ini As String
Dim Cta_Fin As String
Dim Codigo_Ini As String
Dim Codigo_Fin As String
Dim SumModulos(20) As Nodo_Arbol

Private Sub CheqTipoPago_Click()
  If CheqTipoPago.value = 1 Then DCTipoPago.Visible = True Else DCTipoPago.Visible = False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
End Sub

Private Sub Command4_Click()
End Sub

Private Sub Command5_Click()
Dim MesNo As Byte
  If OpcD.value Then
      TextoValido TxtPresMes, True
      MesNo = CMeses.ListIndex + 1
      If MesNo <= 0 Then MesNo = 1
      Codigo1 = CambioCodigoCta(MBoxCta.Text)
      Insertar_SubCtas MesNo, CMeses.Text, Codigo1, Codigo2, Val(CCur(TxtPresMes))
  End If
  FrmPresupuesto.Visible = False
  TVCatalogo.SetFocus
End Sub

Private Sub Command6_Click()
 FrmPresupuesto.Visible = False
 TVCatalogo.SetFocus
End Sub

Private Sub DGGastos_DblClick()
   Codigo2 = Ninguno
   Codigo1 = CambioCodigoCta(MBoxCta.Text)
   If AdoGastos.Recordset.RecordCount > 0 Then Codigo2 = AdoGastos.Recordset.Fields("Codigo")
   SQL1 = "SELECT Codigo,Mes,Presupuesto " _
        & "FROM Trans_Presupuestos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Cta = '" & Codigo1 & "' " _
        & "AND Codigo = '" & Codigo2 & "' " _
        & "ORDER BY Codigo,Mes_No "
   SelectDataGrid DGGastos1, AdoGastos1, SQL1
End Sub

Private Sub DGGastos_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyInsert And CtrlDown Then
     FrmPresupuesto.Top = SSTab1.Top
     FrmPresupuesto.Left = SSTab1.Left
     FrmPresupuesto.Visible = True
     Codigo2 = Ninguno
     If AdoGastos.Recordset.RecordCount > 0 Then Codigo2 = AdoGastos.Recordset.Fields("Codigo")
     CMeses.ListIndex = 0
     CMeses.SetFocus
  End If
End Sub

Private Sub Form_Activate()
Dim CodigoCtas() As String
    
    DGGastos.Visible = False
    SumModulos(0).Item_Nodo = "General/Normal"
    SumModulos(0).Codigo_Aux = "N"
    
    SumModulos(1).Item_Nodo = "Cuenta de Caja"
    SumModulos(1).Codigo_Aux = CtaCaja
    
    SumModulos(2).Item_Nodo = "Cuenta de Bancos"
    SumModulos(2).Codigo_Aux = CtaBancos
    
    SumModulos(3).Item_Nodo = "Modulo de CxC"
    SumModulos(3).Codigo_Aux = "C"
    
    SumModulos(4).Item_Nodo = "Modulo de CxP"
    SumModulos(4).Codigo_Aux = "P"
    
    SumModulos(5).Item_Nodo = "Modulo de Ingresos"
    SumModulos(5).Codigo_Aux = "I"
    
    SumModulos(6).Item_Nodo = "Modulo de Gastos"
    SumModulos(6).Codigo_Aux = "G"
    
    SumModulos(7).Item_Nodo = "CxC Sin Submódulo"
    SumModulos(7).Codigo_Aux = "CS"
    
    SumModulos(8).Item_Nodo = "CxP Sin Submódulo"
    SumModulos(8).Codigo_Aux = "PS"
    
    SumModulos(9).Item_Nodo = "Retención en la Fuente"
    SumModulos(9).Codigo_Aux = "RF"
    
    SumModulos(10).Item_Nodo = "Retención del I.V.A. Servicio"
    SumModulos(10).Codigo_Aux = "RI"
    
    SumModulos(11).Item_Nodo = "Retención del I.V.A. Bienes"
    SumModulos(11).Codigo_Aux = "RB"
    
    SumModulos(12).Item_Nodo = "Crédito Retencion en la Fuente"
    SumModulos(12).Codigo_Aux = "CF"
    
    SumModulos(13).Item_Nodo = "Crédito Retencion del I.V.A. Servicio"
    SumModulos(13).Codigo_Aux = "CI"
    
    SumModulos(14).Item_Nodo = "Crédito Retencion del I.V.A. Bienes"
    SumModulos(14).Codigo_Aux = "CB"
    
    SumModulos(15).Item_Nodo = "Caja Cheques Posfechados"
    SumModulos(15).Codigo_Aux = "CP"
    
    SumModulos(16).Item_Nodo = "Modulo de Primas"
    SumModulos(16).Codigo_Aux = "PM"
    
    SumModulos(17).Item_Nodo = "Modulo de Inventario"
    SumModulos(17).Codigo_Aux = "RP"
    
    SumModulos(18).Item_Nodo = "Opcion Tarjeta de Credito"
    SumModulos(18).Codigo_Aux = "TJ"
    
    SumModulos(19).Item_Nodo = "Modulo Centro de Costos"
    SumModulos(19).Codigo_Aux = "CC"
    
    For I = 0 To UBound(SumModulos) - 1
        LstSubMod.AddItem SumModulos(I).Item_Nodo
    Next I
   
   'Verificamos Nuevas cuentas en proyectos si fuera el caso
    If ConSucursal Then
       Cadena = ""
       sSQL = "SELECT * " _
            & "FROM Catalogo_Cuentas " _
            & "WHERE Item = '" & NumEmpresa & "' " _
            & "AND Periodo = '" & Periodo_Contable & "' " _
            & "ORDER BY Codigo "
       SelectAdodc AdoCta, sSQL
       
       sSQL = "SELECT Codigo,Cuenta,DG,TC " _
            & "FROM Catalogo_Cuentas " _
            & "WHERE Periodo = '" & Periodo_Contable & "' " _
            & "AND Item NOT IN ('" & NumEmpresa & "','000') " _
            & "GROUP BY Codigo,Cuenta,DG,TC " _
            & "ORDER BY Codigo "
       SelectAdodc AdoCtas, sSQL
       'MsgBox AdoAux.Recordset.RecordCount & vbCrLf & AdoEmp000.Recordset.RecordCount
       If AdoCta.Recordset.RecordCount > 0 Then
          With AdoCtas.Recordset
           If .RecordCount > 0 Then
               Do While Not .EOF
                  Codigo = .Fields("Codigo")
                  AdoCta.Recordset.MoveFirst
                  AdoCta.Recordset.Find ("Codigo = '" & Codigo & "' ")
                  If AdoCta.Recordset.EOF Then
                     SetAdoAddNew "Catalogo_Cuentas"
                     SetAdoFields "Item", NumEmpresa
                     SetAdoFields "Periodo", Periodo_Contable
                     SetAdoFields "Codigo", .Fields("Codigo")
                     SetAdoFields "Cuenta", .Fields("Cuenta")
                     SetAdoFields "DG", .Fields("DG")
                     SetAdoFields "TC", .Fields("TC")
                     SetAdoUpdate
                     sSQL = "SELECT * " _
                          & "FROM Catalogo_Cuentas " _
                          & "WHERE Item = '" & NumEmpresa & "' " _
                          & "AND Periodo = '" & Periodo_Contable & "' " _
                          & "ORDER BY Codigo "
                     SelectAdodc AdoCta, sSQL
                     Cadena = Cadena & "Empresa: " & NumEmpresa & ", Cta = " & .Fields("Codigo") & ", Detalle = " & .Fields("Cuenta") & vbCrLf
                  End If
                 .MoveNext
               Loop
           End If
          End With
       End If
       If Cadena <> "" Then MsgBox "CUENTAS INSERTADAS: " & vbCrLf & Cadena
    Else
       'Verificamos si existe Catalogo de Cuenta en la empresa seleccionada
        sSQL = "SELECT * " _
             & "FROM Catalogo_Cuentas " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY Codigo "
        SelectAdodc AdoCta, sSQL
        
        sSQL = "SELECT * " _
             & "FROM Catalogo_Cuentas " _
             & "WHERE Item = '000' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "ORDER BY Codigo "
        SelectAdodc AdoCtas, sSQL
       'MsgBox AdoAux.Recordset.RecordCount & vbCrLf & AdoEmp000.Recordset.RecordCount
        If AdoCta.Recordset.RecordCount <= 0 Then
           With AdoCtas.Recordset
            If .RecordCount > 0 Then
                Do While Not .EOF
                   SetAdoAddNew "Catalogo_Cuentas"
                   SetAdoFields "Item", NumEmpresa
                   SetAdoFields "Periodo", Periodo_Contable
                   SetAdoFields "Codigo", .Fields("Codigo")
                   SetAdoFields "Cuenta", .Fields("Cuenta")
                   SetAdoFields "DG", .Fields("DG")
                   SetAdoFields "TC", .Fields("TC")
                   SetAdoUpdate
                  .MoveNext
                Loop
            End If
           End With
        End If
    End If
    
    sSQL = "UPDATE Catalogo_Cuentas " _
         & "SET Cta_Acreditar = '0' " _
         & "WHERE Cta_Acreditar = '.' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' "
    Conectar_Ado_Execute sSQL

  Si_No = False
  sSQL = "SELECT Item,Codigo " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE DG = 'D' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND ISNUMERIC(MidStrg(Codigo,1,1)) <> 0 " _
       & "ORDER BY Codigo "
  SelectAdodc AdoPresupuestos, sSQL
  If AdoPresupuestos.Recordset.RecordCount > 0 Then
     sSQL = "SELECT * " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND ISNUMERIC(MidStrg(Codigo,1,1)) <> 0 " _
          & "ORDER BY Codigo "
     SelectAdodc AdoCta, sSQL
     If AdoCta.Recordset.RecordCount > 0 Then
        ReDim CodigoCtas(AdoCta.Recordset.RecordCount + 1) As String
        For I = 0 To AdoCta.Recordset.RecordCount
            CodigoCtas(I) = ""
        Next I
     End If
     Contador = 0
     Do While Not AdoPresupuestos.Recordset.EOF
        Codigo = AdoPresupuestos.Recordset.Fields("Codigo")
        Cta_Sup = CodigoCuentaSup(Codigo)
        With AdoCta.Recordset
         If .RecordCount > 0 Then
             Do While (Cta_Sup <> "0")
               .MoveFirst
               .Find ("Codigo Like '" & Cta_Sup & "' ")
                If Not .EOF And Cta_Sup <> "0" Then
                   Cta_Sup = CodigoCuentaSup(Cta_Sup)
                Else
                   Si_No = True: Evaluar = True
                   For I = 0 To AdoCta.Recordset.RecordCount
                       If CodigoCtas(I) = Cta_Sup Then Evaluar = False
                   Next I
                   If Evaluar Then
                      SetAdoAddNew "Catalogo_Cuentas"
                      SetAdoFields "Item", NumEmpresa
                      SetAdoFields "Codigo", Cta_Sup
                      SetAdoFields "Cuenta", "NINGUNA CUENTA"
                      SetAdoFields "Periodo", Periodo_Contable
                      SetAdoFields "DG", "G"
                      SetAdoFields "TC", "N"
                      SetAdoUpdate
                      CodigoCtas(Contador) = Cta_Sup
                      Contador = Contador + 1
                   End If
                   Cta_Sup = CodigoCuentaSup(Cta_Sup)
                End If
             Loop
         End If
        End With
        AdoPresupuestos.Recordset.MoveNext
     Loop
  End If
  If Si_No Then
     Cadena = vbCrLf
     For I = 0 To Contador
         Cadena = Cadena & CodigoCtas(I) & vbCrLf
     Next I
     MsgBox "Los siguientes Codigos no se han creado: " & vbCrLf _
            & Cadena & "ADVERTENCIA: REVIZAR."
  End If
  
  sSQL = "UPDATE Catalogo_Cuentas " _
       & "SET Cuenta = UCaseStrg(Cuenta), TC = 'N' " _
       & "WHERE DG = 'G' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  Conectar_Ado_Execute sSQL
  
  SQL1 = "SELECT * " _
       & "FROM Trans_Presupuestos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY Cta,Codigo "
  SelectAdodc AdoPresupuestos, SQL1
    
'  SQL1 = "UPDATE Catalogo_SubCtas " _
'       & "SET Presupuesto = 0 " _
'       & "WHERE TC = 'G' " _
'       & "AND Item = '" & NumEmpresa & "' " _
'       & "AND Periodo = '" & Periodo_Contable & "' "
'  Conectar_Ado_Execute SQL1
  
  sSQL = "SELECT (Codigo & ' ' & Descripcion) As CTipoPago, Codigo " _
       & "FROM Tabla_Referenciales_SRI " _
       & "WHERE Tipo_Referencia = 'FORMA DE PAGO' " _
       & "AND Codigo IN ('01','16','17','18','19','20','21') " _
       & "ORDER BY Codigo "
  SelectDBCombo DCTipoPago, AdoTipoPago, sSQL, "CTipoPago"
  
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND ISNUMERIC(MidStrg(Codigo,1,1)) <> 0 " _
       & "ORDER BY Codigo "
  SelectAdodc AdoCta, sSQL
  FormatoMaskCta MBoxCta
  FormatoMaskCta MBoxCtaAcreditar
  RatonReloj
  With AdoCta.Recordset
   If .RecordCount > 0 Then
       .MoveFirst
        Do While Not .EOF
            
           If Len(.Fields("Codigo")) = 1 Then
              Codigo = "C" & .Fields("Codigo")
              Cta_Sup = .Fields("Codigo")
              Cuenta = .Fields("Codigo") & " - " & .Fields("Cuenta")
              AddNewCta .Fields("TC")
           Else
              Codigo = "C" & .Fields("Codigo")
              Cta_Sup = "C" & CodigoCuentaSup(.Fields("Codigo"))
              Cuenta = .Fields("Codigo") & " - " & .Fields("Cuenta")
              If .Fields("DG") = "G" Then
                  AddNewCta "DG"
              Else
                  AddNewCta .Fields("TC")
              End If
           End If
          .MoveNext
        Loop
    End If
   End With
   
   CMeses.Clear
   CMeses.AddItem "Enero"
   CMeses.AddItem "Febrero"
   CMeses.AddItem "Marzo"
   CMeses.AddItem "Abril"
   CMeses.AddItem "Mayo"
   CMeses.AddItem "Junio"
   CMeses.AddItem "Julio"
   CMeses.AddItem "Agosto"
   CMeses.AddItem "Septiembre"
   CMeses.AddItem "Octubre"
   CMeses.AddItem "Noviembre"
   CMeses.AddItem "Diciembre"
   CMeses.AddItem "Todos"
   If Modo_Educativo Then Command3.Enabled = False
   Select Case CodigoUsuario
     Case "ACCESO01", "ACCESO02", "ACCESO03", "ACCESO04", "ACCESO05", "0702164179"
          Command3.Enabled = True
          Command4.Enabled = True
   End Select
   Command1.SetFocus
   
   If Bloquear_Control Then
      Command1.Enabled = False
      Command3.Enabled = False
      Command4.Enabled = False
   End If
   
   RatonNormal
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoCta
  ConectarAdodc AdoCtas
  ConectarAdodc AdoGastos
  ConectarAdodc AdoGastos1
  ConectarAdodc AdoTipoPago
  ConectarAdodc AdoPresupuesto
  ConectarAdodc AdoPresupuestos
End Sub

Private Sub LstSubMod_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub LstSubMod_LostFocus()
  Label10.Visible = False
  MBoxCtaAcreditar.Visible = False
  For I = 0 To UBound(SumModulos) - 1
      Select Case SumModulos(I).Codigo_Aux
       Case "G", "I"
            Codigo1 = CambioCodigoCta(MBoxCta.Text)
            SQL1 = "UPDATE Catalogo_SubCtas " _
                 & "SET Presupuesto = 0 " _
                 & "WHERE TC = '" & SumModulos(I).Codigo_Aux & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' "
            Conectar_Ado_Execute SQL1
            
            SQL1 = "SELECT Codigo,Detalle,Presupuesto " _
                 & "FROM Catalogo_SubCtas " _
                 & "WHERE TC = '" & SumModulos(I).Codigo_Aux & "' " _
                 & "AND Item = '" & NumEmpresa & "' " _
                 & "AND Periodo = '" & Periodo_Contable & "' " _
                 & "ORDER BY Codigo "
            SelectDataGrid DGGastos, AdoGastos, SQL1
            DGGastos.Visible = True
       Case "TJ"
            Label10.Visible = True
            MBoxCtaAcreditar.Visible = True
      End Select
  Next I
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
       & "FROM Catalogo_Cuentas " _
       & "WHERE Codigo = '" & Codigo & "' " _
       & "AND Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' "
  SelectAdodc AdoCtas, sSQL, False
  If (AdoCtas.Recordset.RecordCount <= 0) And (Len(Codigo) > 1) Then
     Cadena = "Warnign: No puede crear este Código," & vbCrLf _
            & "no existe Cuenta Superior "
     MsgBox Cadena
     MBoxCta.SetFocus
  Else
     LabelCtaSup.Caption = CambioCodigoCtaSup(CambioCodigoCta(MBoxCta.Text))
     Codigos = CambioCodigoCta(MBoxCta.Text)
     sSQL = "SELECT Codigo " _
          & "FROM Catalogo_Cuentas " _
          & "WHERE Codigo = '" & Codigos & "' " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' "
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
  DGGastos.Visible = False
  With AdoCta.Recordset
   If .RecordCount > 0 Then
       TipoSubCta = .Fields("TC")
       If .Fields("ME") Then CheqUS.value = 1 Else CheqUS.value = 0
       If .Fields("Listar") Then CheqFE.value = 1 Else CheqFE.value = 0
       If .Fields("Mod_Gastos") Then CheqModGastos.value = 1 Else CheqModGastos.value = 0
       For I = 0 To UBound(SumModulos) - 1
          If SumModulos(I).Codigo_Aux = .Fields("TC") Then LstSubMod.Text = SumModulos(I).Item_Nodo
       Next I
       Cadena = .Fields("DG")
       If Cadena = "D" Then OpcD.value = True Else OpcG.value = True
       MBoxCta.Text = FormatoCodigoCta(.Fields("Codigo"))
       MBoxCtaAcreditar = FormatoCodigoCta(.Fields("Cta_Acreditar"))
       LabelCtaSup.Caption = CodigoCuentaSup(.Fields("Codigo"))
       LabelNumero.Caption = .Fields("Clave")
       TextConcepto.Text = .Fields("Cuenta")
       LabelTipoCta.Caption = TiposCtaStrg(.Fields("Codigo"))
       TextPresupuesto.Text = .Fields("Presupuesto")
       TxtCodExt.Text = .Fields("Codigo_Ext")
       If Val(.Fields("Tipo_Pago")) >= 1 Then
          AdoTipoPago.Recordset.MoveFirst
          AdoTipoPago.Recordset.Find ("Codigo = '" & .Fields("Tipo_Pago") & "' ")
          If Not AdoTipoPago.Recordset.EOF Then DCTipoPago = AdoTipoPago.Recordset.Fields("CTipoPago")
          CheqTipoPago.value = 1
          DCTipoPago.Visible = True
       Else
          CheqTipoPago.value = 0
          DCTipoPago.Visible = False
       End If
       If .Fields("I_E_Emp") = Ninguno Then
           OpcNoAplica.value = True
           CheqConIESS.value = 0
       ElseIf .Fields("I_E_Emp") = "I" Then
           OpcIEmp.value = True
           If .Fields("Con_IESS") Then CheqConIESS.value = 1 Else CheqConIESS.value = 0
       Else
          OpcEEmp.value = True
          CheqConIESS.value = 0
       End If
       Nuevo = False
   Else
      Nuevo = True
   End If
   End With
   If OpcGas Or OpcI Then
      SQL1 = "UPDATE Catalogo_SubCtas " _
           & "SET Presupuesto = 0 " _
           & "WHERE TC = '" & TipoSubCta & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' "
      Conectar_Ado_Execute SQL1

      Codigo1 = CambioCodigoCta(MBoxCta.Text)
      If SQL_Server Then
         sSQL = "UPDATE Catalogo_SubCtas " _
              & "SET Presupuesto = P.Presupuesto " _
              & "FROM Catalogo_SubCtas As B,Trans_Presupuestos As P "
      Else
         sSQL = "UPDATE Catalogo_SubCtas As B,Trans_Presupuestos As P " _
              & "SET B.Presupuesto = P.Presupuesto "
      End If
      sSQL = sSQL & "WHERE P.Cta = '" & Codigo1 & "' " _
           & "AND B.Item = '" & NumEmpresa & "' " _
           & "AND B.Periodo = '" & Periodo_Contable & "' " _
           & "AND B.Codigo = P.Codigo " _
           & "AND B.Periodo = P.Periodo " _
           & "AND B.Item = P.Item "
      Conectar_Ado_Execute sSQL
      
      SQL1 = "SELECT Codigo,Detalle,Presupuesto,Periodo,Item " _
           & "FROM Catalogo_SubCtas " _
           & "WHERE TC = '" & TipoSubCta & "' " _
           & "AND Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "ORDER BY Codigo "
      SelectDataGrid DGGastos, AdoGastos, SQL1
      DGGastos.Visible = True
   ElseIf OpcCxC Or OpcCxP Then
      Codigo1 = CambioCodigoCta(MBoxCta.Text)
      SQL1 = "SELECT C.Cliente,CCP.Codigo,CCP.Periodo,CCP.Item " _
           & "FROM Catalogo_CxCxP As CCP, Clientes As C " _
           & "WHERE CCP.Cta = '" & Codigo1 & "' " _
           & "AND CCP.Item = '" & NumEmpresa & "' " _
           & "AND CCP.Periodo = '" & Periodo_Contable & "' " _
           & "AND CCP.Codigo = C.Codigo " _
           & "ORDER BY C.Cliente "
      SelectDataGrid DGGastos, AdoGastos, SQL1
      DGGastos.Visible = True
   Else
      DGGastos.Visible = False
   End If
   Label6.Visible = True
   Codigo1 = CambioCodigoCta(MBoxCta.Text)
   SQL1 = "SELECT Mes,Presupuesto " _
        & "FROM Trans_Presupuestos " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Cta = '" & Codigo1 & "' " _
        & "AND Codigo = '" & Ninguno & "' " _
        & "ORDER BY Mes_No "
   SelectDataGrid DGPresupuesto, AdoPresupuesto, SQL1
   Total = 0
   With AdoPresupuesto.Recordset
    If .RecordCount > 0 Then
        Do While Not .EOF
           Total = Total + .Fields("Presupuesto")
          .MoveNext
        Loop
    End If
   End With
   TextPresupuesto = Format(Total, "#,##0.00")
End Sub

Public Sub GrabarCta(NuevaCta As Boolean)
  If OpcG.value Then TipoDoc = "G" Else TipoDoc = "D"
  If CheqTipoPago.value = 1 Then FA.Tipo_Pago = SinEspaciosIzq(DCTipoPago) Else FA.Tipo_Pago = "00"
  NuevaCta = False
  TextoValido TextPresupuesto
  If LabelCtaSup.Caption = "" Then LabelCtaSup.Caption = "0"
  Numero = 0
  TipoCta = "N"
  For I = 0 To UBound(SumModulos) - 1
   If LstSubMod.Text = SumModulos(I).Item_Nodo Then
      TipoCta = SumModulos(I).Codigo_Aux
      'MsgBox SumModulos(I).Codigo_Aux
   End If
  Next I
  If TipoDoc = "G" Then
     TextoValido TextConcepto, , True
  Else
     TextoValido TextConcepto
  End If
  Codigo1 = CambioCodigoCta(MBoxCta.Text)
  Codigo = "C" & Codigo1
  Cta_Sup = "C" & CodigoCuentaSup(Codigo1)
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
             Numero = .Fields("Clave")
             If OpcD.value And Numero = 0 Then
                Numero = ReadSetDataNum("Numero Cuenta", True, True)
             End If
          Else
            .AddNew
            .Fields("Codigo") = Codigo1
             If OpcD.value Then
                Numero = ReadSetDataNum("Numero Cuenta", True, True)
             End If
             AddNewCta TipoCta
             NuevaCta = True
          End If
      Else
         .AddNew
         .Fields("Codigo") = Codigo1
          If OpcD.value Then
             Numero = ReadSetDataNum("Numero Cuenta", True, True)
          End If
          If OpcG.value Then AddNewCta "DG" Else AddNewCta TipoCta
      End If
     ' MsgBox TipoCta
     .Fields("Clave") = Numero
     .Fields("DG") = TipoDoc
     .Fields("TC") = TipoCta
     .Fields("ME") = CheqUS.value
     .Fields("Listar") = CheqFE.value
     .Fields("Mod_Gastos") = CheqModGastos.value
     .Fields("Cuenta") = TextConcepto.Text
     .Fields("Presupuesto") = CCur(TextPresupuesto.Text)
     .Fields("Procesado") = vbTrue
     .Fields("Periodo") = Periodo_Contable
     .Fields("Item") = NumEmpresa
     .Fields("Codigo_Ext") = TxtCodExt
     .Fields("Cta_Acreditar") = CambioCodigoCta(MBoxCtaAcreditar)
     .Fields("Tipo_Pago") = FA.Tipo_Pago
      If OpcNoAplica.value Then
        .Fields("I_E_Emp") = Ninguno
        .Fields("Con_IESS") = False
        .Fields("Cod_Rol_Pago") = Ninguno
      Else
        .Fields("Cod_Rol_Pago") = Rubro_Rol_Pago(TextConcepto)
         If OpcIEmp.value Then
           .Fields("I_E_Emp") = "I"
            If CheqConIESS.value <> 0 Then .Fields("Con_IESS") = True Else .Fields("Con_IESS") = False
         Else
           .Fields("I_E_Emp") = "E"
         End If
      End If
     .Update
      UpdateCta TipoCta
     End With
  End If
  If OpcCxCP Then
     Mensajes = "Ingrese la Cuenta de Interes:"
     Titulo = "Cuenta de Interés para el Prestamo"
     TextoCheque = InputBox(Mensajes, Titulo, "")
     
     If TextoCheque = "" Then TextoCheque = "1"
     MsgBox TextoCheque
     SQL1 = "SELECT * " _
          & "FROM Ctas_Proceso " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "ORDER BY T_No "
     SelectDataGrid DGGastos, AdoGastos, SQL1
     If AdoGastos.Recordset.RecordCount > 0 Then
        AdoGastos.Recordset.MoveLast
        Contador = AdoGastos.Recordset.Fields("T_No") + 1
        Si_No = True
        Do While Not AdoGastos.Recordset.EOF And Si_No
           If AdoGastos.Recordset.Fields("Detalle") = Codigo1 Then Si_No = False
           AdoGastos.Recordset.MoveNext
        Loop
        If Si_No Then
           AdoGastos.Recordset.AddNew
           AdoGastos.Recordset.Fields("DC") = "C"
           AdoGastos.Recordset.Fields("T_No") = Contador
           AdoGastos.Recordset.Fields("Detalle") = Codigo1
           AdoGastos.Recordset.Fields("Item") = NumEmpresa
        End If
        AdoGastos.Recordset.Fields("Codigo") = TextoCheque
        AdoGastos.Recordset.Fields("Lst") = False
        AdoGastos.Recordset.Update
     End If
  End If
  If OpcGas Or OpcI Then
     If AdoGastos.Recordset.RecordCount > 0 Then
        AdoGastos.Recordset.MoveFirst
        Codigo1 = CambioCodigoCta(MBoxCta)
        SQL1 = "DELETE * " _
             & "FROM Trans_Presupuestos " _
             & "WHERE Cta = '" & Codigo1 & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' "
        Conectar_Ado_Execute SQL1
        Do While Not AdoGastos.Recordset.EOF
           Valor = AdoGastos.Recordset.Fields("Presupuesto")
           Codigo = AdoGastos.Recordset.Fields("Codigo")
           If Valor >= 0 Then
              AdoPresupuestos.Recordset.AddNew
              AdoPresupuestos.Recordset.Fields("Cta") = Codigo1
              AdoPresupuestos.Recordset.Fields("Codigo") = Codigo
              AdoPresupuestos.Recordset.Fields("Presupuesto") = Valor
              AdoPresupuestos.Recordset.Fields("Item") = NumEmpresa
              AdoPresupuestos.Recordset.Fields("Periodo") = Periodo_Contable
              AdoPresupuestos.Recordset.Update
           End If
           AdoGastos.Recordset.MoveNext
        Loop
     End If
  End If
  If NuevaCta Then
     Control_Procesos Normal, "Nuva Cuenta: " & Codigo1 & " - " & TextConcepto.Text
  Else
     Control_Procesos Normal, "Modificacion de Cuenta: " & Codigo1 & " - " & TextConcepto.Text
  End If
  sSQL = "SELECT * " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND MidStrg(Codigo,1,1) <> 'x' " _
       & "ORDER BY Codigo "
  SelectAdodc AdoCta, sSQL
  IE = TVCatalogo.SelectedItem.Index
  If NuevaCta = False Then TVCatalogo.Nodes(IE).Text = Codigo1 & " - " & TextConcepto.Text
  TVCatalogo.Refresh
  Label6.Visible = True
  Nuevo = False
End Sub

Public Sub NuevaCta()
  OpcNor.value = True
  LabelNumero.Caption = "0"
  LabelNumero.Caption = ""
  TextConcepto.Text = ""
  TextPresupuesto.Text = ""
  LabelCtaSup.Caption = ""
  MBoxCta.Text = LimpiarCtas
  Nuevo = True
  MBoxCta.SetFocus
End Sub

Private Sub TextPresupuesto_GotFocus()
  MarcarTexto TextPresupuesto
End Sub

Private Sub TextPresupuesto_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  If KeyCode = vbKeyInsert And CtrlDown Then
     FrmPresupuesto.Top = TextPresupuesto.Top
     FrmPresupuesto.Left = TextPresupuesto.Left
     FrmPresupuesto.Visible = True
     Codigo2 = Ninguno
     CMeses.ListIndex = 0
     CMeses.SetFocus
  End If
End Sub

Private Sub TextPresupuesto_LostFocus()
  TextoValido TextPresupuesto, True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   'MsgBox Button.key
    Select Case Button.key
           Case "Salir"
                Unload FCatalogo_Cuentas
           Case "Grabar"
                If Nuevo Then GrabarCta (True) Else GrabarCta (False)
           Case "Copiar"
                If ClaveSupervisor Then
                   RatonReloj
                   Si_No = False
                   FCopyCat.Show 1
                End If
           Case "Cambiar"
                If ClaveSupervisor Then
                   RatonReloj
                   Producto = "Catalogo"
                   If OpcD.value Then
                      Codigo1 = CambioCodigoCta(MBoxCta.Text)
                      Codigo3 = Codigo1 & " - " & TextConcepto
                      FChangeCta.Show 1
                   Else
                      RatonNormal
                      MsgBox "Solo puede cambiar Cuentas de Detalle"
                   End If
                End If
    End Select
End Sub

Private Sub TVCatalogo_DblClick()
  SiguienteControl
End Sub

Private Sub TVCatalogo_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
  PresionoEnter KeyCode
  Codigo1 = SinEspaciosIzq(TVCatalogo.SelectedItem)
  Codigo2 = MidStrg(TVCatalogo.SelectedItem, Len(Codigo1) + 1, Len(TVCatalogo.SelectedItem))
  If CtrlDown And KeyCode = vbKeyI Then
     Codigo_Ini = Codigo1
     Cta_Ini = Codigo2
  End If
  If CtrlDown And KeyCode = vbKeyU Then
     Codigo_Fin = Codigo1
     Cta_Fin = Codigo2
  End If
  If CtrlDown And KeyCode = vbKeyDelete Then EliminarCtaGrupo
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
     TVCatalogo.Nodes.Add , , Codigo, Cuenta, ImageList1.ListImages(9).key, ImageList1.ListImages(9).key
  Else
     Select Case TipoTC
       Case "DG": IE = 9
       Case "RF": IE = 11
       Case "CF": IE = 11
       Case "RI": IE = 21
       Case "RS": IE = 21
       Case "RP": IE = 11
       Case "CI": IE = 11
       Case "CB": IE = 11
       Case "C": IE = 12
       Case "P": IE = 13
       Case "I": IE = 14
       Case "G": IE = 15
       Case "CJ": IE = 16
       Case "BA": IE = 17
       Case "CS": IE = 18
       Case "PS": IE = 19
       Case "CP": IE = 24
       Case "PM": IE = 22
       Case "TJ": IE = 23
       Case "CC": IE = 20
       Case Else: IE = 10
     End Select
     TVCatalogo.Nodes.Add Cta_Sup, tvwChild, Codigo, Cuenta, ImageList1.ListImages(IE).key, ImageList1.ListImages(IE).key
  End If
 'MsgBox MidStrg(Codigo, 2, Len(Codigo)) & vbCrLf & Codigo
  TVCatalogo.Tag = MidStrg(Codigo, 2, Len(Codigo))
End Sub

Public Sub UpdateCta(TipoTC As String)
 ' TVCatalogo.SelectedItem = Cuenta
  Select Case TipoTC
    Case "DG": IE = 9
    Case "RF": IE = 11
    Case "CF": IE = 11
    Case "RI": IE = 21
    Case "RS": IE = 21
    Case "RP": IE = 11
    Case "CI": IE = 11
    Case "CB": IE = 11
    Case "C": IE = 12
    Case "P": IE = 13
    Case "I": IE = 14
    Case "G": IE = 15
    Case "CJ": IE = 16
    Case "BA": IE = 17
    Case "CS": IE = 18
    Case "PS": IE = 19
    Case "CP": IE = 24
    Case "PM": IE = 22
    Case "TJ": IE = 23
    Case "CC": IE = 20
    Case Else: IE = 10
  End Select
'  nodX.Image = ImageList1.ListImages(IE).key
'  nodX.SelectedImage = ImageList1.ListImages(IE).key
End Sub

Public Sub EliminarCta()
  Codigo1 = CambioCodigoCta(MBoxCta)
  Cadena = SinEspaciosIzq(TVCatalogo.SelectedItem)
  Codigo2 = TrimStrg(MidStrg(TVCatalogo.SelectedItem, Len(Cadena) + 1, Len(TVCatalogo.SelectedItem)))
  With AdoCta.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Codigo like '" & Cadena & "' ")
       If Not .EOF Then
          sSQL = "SELECT Cta,Count(Cta) As Cant_Cta " _
               & "FROM Transacciones " _
               & "WHERE MidStrg(Cta,1," & Len(Cadena) & ") = '" & Cadena & "' " _
               & "AND Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "GROUP BY Cta " _
               & "ORDER BY Cta "
          SelectAdodc AdoCtas, sSQL, False
          If AdoCtas.Recordset.RecordCount > 0 Then
             Mensajes = "ADVERTENCIA:" & vbCrLf & vbCrLf _
                      & "No se puede eliminar esta(s) Cuenta(s): " & vbCrLf
             Do While Not AdoCtas.Recordset.EOF
                Mensajes = Mensajes & AdoCtas.Recordset.Fields("Cta") & " Cantidad de Movimientos: " & AdoCtas.Recordset.Fields("Cant_Cta") & vbCrLf
                AdoCtas.Recordset.MoveNext
             Loop
             Mensajes = Mensajes & vbCrLf & "porque tiene(n) novimiento(s)."
             MsgBox Mensajes
          Else
             Mensajes = "Esta seguro que desea eliminar la Cuenta:" & vbCrLf & vbCrLf _
                      & Cadena & ": " & Codigo2 & vbCrLf & vbCrLf _
                      & "y sus grupos "
             Titulo = "Pregunta de Eliminacion"
             If BoxMensaje = vbYes Then
                Cadena1 = SinEspaciosIzq(TVCatalogo.SelectedItem)
                For I = TVCatalogo.Nodes.Count To 1 Step -1
                    CodigoC = MidStrg(TVCatalogo.Nodes(I).key, 2, Len(TVCatalogo.Nodes(I).key))
                    If Cadena1 = MidStrg(CodigoC, 1, Len(Cadena1)) Then
                       SQL1 = "DELETE * " _
                            & "FROM Trans_Presupuestos " _
                            & "WHERE Cta = '" & CodigoC & "' " _
                            & "AND Item = '" & NumEmpresa & "' " _
                            & "AND Periodo = '" & Periodo_Contable & "' "
                       Conectar_Ado_Execute SQL1
                      'MsgBox Cadena1 & vbCrLf & CodigoC & vbCrLf & TVCatalogo.Nodes(I).key & vbCrLf & SQL1
                      .MoveFirst
                      .Find ("Codigo = '" & CodigoC & "' ")
                       If Not .EOF Then
                         .Delete
                          TVCatalogo.Nodes.Remove I
                       End If
                    End If
                Next I
             End If
          End If
       End If
   End If
  End With
End Sub

Public Sub EliminarCtaGrupo()
  With AdoCta.Recordset
   If .RecordCount > 0 Then
        sSQL = "SELECT Cta,Count(Cta) As Cant_Cta " _
             & "FROM Transacciones " _
             & "WHERE Cta BETWEEN '" & Codigo_Ini & "' and '" & codigofin & "' " _
             & "AND Item = '" & NumEmpresa & "' " _
             & "AND Periodo = '" & Periodo_Contable & "' " _
             & "GROUP BY Cta " _
             & "ORDER BY Cta "
        SelectAdodc AdoCtas, sSQL, False
        If AdoCtas.Recordset.RecordCount > 0 Then
           Mensajes = "ADVERTENCIA:" & vbCrLf & vbCrLf _
                    & "No se puede eliminar esta(s) Cuenta(s): " & vbCrLf
           Do While Not AdoCtas.Recordset.EOF
              Mensajes = Mensajes & AdoCtas.Recordset.Fields("Cta") & " Cantidad de Movimientos: " & AdoCtas.Recordset.Fields("Cant_Cta") & vbCrLf
              AdoCtas.Recordset.MoveNext
           Loop
           Mensajes = Mensajes & vbCrLf & "porque tiene(n) novimiento(s)."
           MsgBox Mensajes
        Else
           Mensajes = "Esta seguro que desea eliminar" & vbCrLf & vbCrLf _
                    & "Desde: " & Codigo_Ini & " hasta " & Codigo_Fin & vbCrLf & vbCrLf _
                    & "y sus grupos "
           Titulo = "Pregunta de Eliminacion"
           If BoxMensaje = vbYes Then
              For I = TVCatalogo.Nodes.Count To 1 Step -1
                  If Codigo_Ini <= TVCatalogo.Nodes(I).Tag And TVCatalogo.Nodes(I).Tag <= Codigo_Fin Then
                     SQL1 = "DELETE * " _
                          & "FROM Trans_Presupuestos " _
                          & "WHERE Cta = '" & TVCatalogo.Nodes(I).Tag & "' " _
                          & "AND Item = '" & NumEmpresa & "' " _
                          & "AND Periodo = '" & Periodo_Contable & "' "
                     Conectar_Ado_Execute SQL1
                    .MoveFirst
                    .Find ("Codigo like '" & TVCatalogo.Nodes(I).Tag & "' ")
                     If Not .EOF Then
                       .Delete
                        TVCatalogo.Nodes.Remove I
                     End If
                  End If
              Next I
           End If
        End If
   End If
  End With
End Sub

Private Sub TxtCodExt_GotFocus()
   MarcarTexto TxtCodExt
End Sub

Private Sub TxtCodExt_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtCodExt_LostFocus()
  TextoValido TxtCodExt
End Sub

Private Sub TxtPresMes_GotFocus()
  MarcarTexto TxtPresMes
End Sub

Private Sub TxtPresMes_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Insertar_SubCtas(MesNo As Byte, Mes As String, Cta As String, CodigoSubCta As String, TValor As Currency)
Dim Id_Mes As Byte
   If Mes = "Todos" Then
      For Id_Mes = 1 To 12
          Mifecha = "01/" & Format(Id_Mes, "00") & "/" & Year(FechaSistema)
          SQL1 = "DELETE * " _
               & "FROM Trans_Presupuestos " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND Cta = '" & Cta & "' " _
               & "AND Codigo = '" & Codigo2 & "' " _
               & "AND Mes_No = #" & BuscarFecha(Mifecha) & "# "
          Conectar_Ado_Execute SQL1
          SetAdoAddNew "Trans_Presupuestos"
          SetAdoFields "Mes_No", Mifecha
          SetAdoFields "Cta", Cta
          SetAdoFields "Mes", UCaseStrg(MidStrg(MesesLetras(CInt(Id_Mes)), 1, 3))
          SetAdoFields "Codigo", CodigoSubCta
          SetAdoFields "Presupuesto", TValor
          SetAdoUpdate
      Next Id_Mes
   Else
      Mifecha = "01/" & Format(MesNo, "00") & "/" & Year(FechaSistema)
      SQL1 = "DELETE * " _
           & "FROM Trans_Presupuestos " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "AND Periodo = '" & Periodo_Contable & "' " _
           & "AND Cta = '" & Cta & "' " _
           & "AND Codigo = '" & Codigo2 & "' " _
           & "AND Mes_No = #" & BuscarFecha(Mifecha) & "# "
      Conectar_Ado_Execute SQL1
      SetAdoAddNew "Trans_Presupuestos"
      SetAdoFields "Mes_No", Mifecha
      SetAdoFields "Cta", Cta
      SetAdoFields "Mes", UCaseStrg(MidStrg(Mes, 1, 3))
      SetAdoFields "Codigo", CodigoSubCta
      SetAdoFields "Presupuesto", TValor
      SetAdoUpdate
   End If
End Sub
