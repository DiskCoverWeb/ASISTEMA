VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FacturasPV 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   15645
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Left            =   19005
      Top             =   210
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   16695
      Picture         =   "FacturaH.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   7770
      Width           =   2115
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   14175
      Picture         =   "FacturaH.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   7770
      Width           =   2220
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   6840
      Left            =   14175
      TabIndex        =   48
      Top             =   840
      Width           =   4635
      Begin VB.TextBox TextCheque 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2520
         MaxLength       =   14
         TabIndex        =   68
         Text            =   "0.00"
         Top             =   5775
         Width           =   2010
      End
      Begin VB.TextBox TextBanco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   105
         MaxLength       =   25
         TabIndex        =   66
         Top             =   5250
         Width           =   4425
      End
      Begin VB.TextBox TextCheqNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   64
         Top             =   4305
         Width           =   2010
      End
      Begin VB.TextBox TxtEfectivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   420
         Left            =   2520
         MaxLength       =   12
         MultiLine       =   -1  'True
         TabIndex        =   60
         Text            =   "FacturaH.frx":0BD4
         Top             =   2835
         Width           =   2010
      End
      Begin MSDataListLib.DataCombo DCBanco 
         Bindings        =   "FacturaH.frx":0BDB
         DataSource      =   "AdoBanco"
         Height          =   420
         Left            =   105
         TabIndex        =   62
         Top             =   3780
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   741
         _Version        =   393216
         Text            =   "Banco"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VALOR BANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   105
         TabIndex        =   67
         Top             =   5775
         Width           =   2430
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NOMBRE DEL BANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   65
         Top             =   4830
         Width           =   4425
      End
      Begin VB.Label LblCambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2520
         TabIndex        =   70
         Top             =   6300
         Width           =   2010
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DOCUMENTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   63
         Top             =   4305
         Width           =   2430
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CUENTA DEL BANCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   61
         Top             =   3360
         Width           =   4425
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CAMBIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   69
         Top             =   6300
         Width           =   2430
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Tarifa 0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   49
         Top             =   210
         Width           =   2430
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Tarifa 12%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   51
         Top             =   735
         Width           =   2430
      End
      Begin VB.Label LabelSubTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2520
         TabIndex        =   50
         Top             =   210
         Width           =   2010
      End
      Begin VB.Label LabelConIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2520
         TabIndex        =   52
         Top             =   735
         Width           =   2010
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " I.V.A. 12%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   105
         TabIndex        =   53
         Top             =   1260
         Width           =   2430
      End
      Begin VB.Label Label26 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Facturado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   55
         Top             =   1785
         Width           =   2430
      End
      Begin VB.Label LabelIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2520
         TabIndex        =   54
         Top             =   1260
         Width           =   2010
      End
      Begin VB.Label LabelTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2520
         TabIndex        =   56
         Top             =   1785
         Width           =   2010
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EFECTIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   105
         TabIndex        =   59
         Top             =   2835
         Width           =   2430
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Fact. (ME)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   57
         Top             =   2310
         Width           =   2430
      End
      Begin VB.Label LabelTotalME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2520
         TabIndex        =   58
         Top             =   2310
         Width           =   2010
      End
   End
   Begin VB.TextBox TxtObservacion 
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
      Left            =   7140
      MaxLength       =   60
      MultiLine       =   -1  'True
      TabIndex        =   47
      Top             =   8400
      Width           =   6945
   End
   Begin VB.TextBox TxtNota 
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
      MaxLength       =   60
      MultiLine       =   -1  'True
      TabIndex        =   45
      Top             =   8400
      Width           =   6945
   End
   Begin VB.TextBox TxtGavetas 
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
      Left            =   7140
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   1260
      Width           =   1170
   End
   Begin VB.Frame FrmRifa 
      Caption         =   "Ingrese el rango de Boletos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   4515
      TabIndex        =   34
      Top             =   3255
      Visible         =   0   'False
      Width           =   2640
      Begin VB.CommandButton Command4 
         Caption         =   "Cancelar"
         Height          =   330
         Left            =   1575
         TabIndex        =   40
         Top             =   1155
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   330
         Left            =   105
         TabIndex        =   39
         Top             =   1155
         Width           =   960
      End
      Begin VB.TextBox TxtRifaH 
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
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1050
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   38
         Text            =   "FacturaH.frx":0BF2
         Top             =   735
         Width           =   1485
      End
      Begin VB.TextBox TxtRifaD 
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
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1050
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "FacturaH.frx":0BF4
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta"
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
         TabIndex        =   37
         Top             =   735
         Width           =   960
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde"
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
         TabIndex        =   35
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.TextBox TextVUnit 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   9765
      MaxLength       =   10
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "FacturaH.frx":0BF6
      Top             =   1995
      Width           =   1170
   End
   Begin VB.TextBox TextCant 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8715
      MaxLength       =   6
      MultiLine       =   -1  'True
      TabIndex        =   27
      Text            =   "FacturaH.frx":0BFB
      Top             =   1995
      Width           =   960
   End
   Begin VB.TextBox TxtDocumentos 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   12600
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   1995
      Width           =   1485
   End
   Begin VB.Frame FrmGrupo 
      BackColor       =   &H00400000&
      Caption         =   "GRUPO DE FACTURACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3375
      Left            =   7770
      TabIndex        =   41
      Top             =   3150
      Visible         =   0   'False
      Width           =   5895
      Begin MSDataListLib.DataList DLGrupo 
         Bindings        =   "FacturaH.frx":0BFD
         DataSource      =   "AdoGrupo"
         Height          =   2940
         Left            =   105
         TabIndex        =   42
         Top             =   315
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   5186
         _Version        =   393216
         BackColor       =   16744576
         ForeColor       =   192
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
   End
   Begin MSDataListLib.DataCombo DCBodega 
      Bindings        =   "FacturaH.frx":0C14
      DataSource      =   "AdoBodega"
      Height          =   360
      Left            =   9660
      TabIndex        =   21
      Top             =   1260
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   192
      ForeColor       =   16777215
      Text            =   "DC"
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
   Begin VB.TextBox TextFacturaNo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   17220
      TabIndex        =   10
      Text            =   "9999999999"
      Top             =   420
      Width           =   1590
   End
   Begin VB.TextBox TextCotiza 
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
      Left            =   1470
      MaxLength       =   11
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   1260
      Width           =   1380
   End
   Begin VB.OptionButton OpcMult 
      Caption         =   "(x)"
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
      Left            =   4410
      TabIndex        =   16
      Top             =   1260
      Value           =   -1  'True
      Width           =   750
   End
   Begin VB.OptionButton OpcDiv 
      Caption         =   "(/)"
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
      TabIndex        =   17
      Top             =   1260
      Width           =   750
   End
   Begin MSDataListLib.DataCombo DCArticulo 
      Bindings        =   "FacturaH.frx":0C2C
      DataSource      =   "AdoArticulo"
      Height          =   315
      Left            =   105
      TabIndex        =   23
      Top             =   1995
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   16711680
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGAsientoF 
      Bindings        =   "FacturaH.frx":0C46
      Height          =   5580
      Left            =   105
      TabIndex        =   43
      Top             =   2415
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   9843
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12648447
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin MSAdodcLib.Adodc AdoAsientoF 
      Height          =   330
      Left            =   525
      Top             =   4515
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "AsientoF"
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
   Begin MSAdodcLib.Adodc AdoLinea 
      Height          =   330
      Left            =   525
      Top             =   4830
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Linea"
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
   Begin MSAdodcLib.Adodc AdoFactura 
      Height          =   330
      Left            =   525
      Top             =   5145
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Factura"
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
   Begin MSAdodcLib.Adodc AdoArticulo 
      Height          =   330
      Left            =   525
      Top             =   4515
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "Articulo"
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
   Begin MSMask.MaskEdBox MBFecha 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Formato de Fecha: DD/MM/AA"
      Top             =   420
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
   Begin MSAdodcLib.Adodc AdoBodega 
      Height          =   330
      Left            =   525
      Top             =   5460
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Bodega"
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
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   525
      Top             =   5775
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Grupo"
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
   Begin MSAdodcLib.Adodc AdoBenef 
      Height          =   330
      Left            =   2520
      Top             =   4515
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Benef"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   2625
      Top             =   4830
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Aux"
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
   Begin MSAdodcLib.Adodc AdoBanco 
      Height          =   330
      Left            =   2625
      Top             =   5145
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Grupo"
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
   Begin MSDataListLib.DataCombo DCCliente 
      Bindings        =   "FacturaH.frx":0C60
      DataSource      =   "AdoBenef"
      Height          =   315
      Left            =   1470
      TabIndex        =   3
      Top             =   420
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "CONSUMIDOR FINAL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCDireccion 
      Bindings        =   "FacturaH.frx":0C77
      DataSource      =   "AdoDireccion"
      Height          =   315
      Left            =   1470
      TabIndex        =   12
      Top             =   840
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "SD"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoDireccion 
      Height          =   330
      Left            =   2625
      Top             =   5775
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Direccion"
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
   Begin VB.Label Label29 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DIRECCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label Label30 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALDO PENDIENTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   13965
      TabIndex        =   6
      Top             =   105
      Width           =   2010
   End
   Begin VB.Label LblSaldoPendiente 
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   13965
      TabIndex        =   7
      Top             =   420
      Width           =   2010
   End
   Begin VB.Label Label24 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OBSERVACION:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7140
      TabIndex        =   46
      Top             =   8085
      Width           =   6945
   End
   Begin VB.Label Label21 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOTA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   44
      Top             =   8085
      Width           =   6945
   End
   Begin VB.Label Label18 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " GAVETAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5985
      TabIndex        =   18
      Top             =   1260
      Width           =   1170
   End
   Begin VB.Label LblRUC 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   12180
      TabIndex        =   5
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CI/RUC/PAS."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   12180
      TabIndex        =   4
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label LblSerie 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999-999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   16065
      TabIndex        =   9
      Top             =   420
      Width           =   1170
   End
   Begin VB.Label LabelVTotal 
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
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   11025
      TabIndex        =   31
      Top             =   1995
      Width           =   1485
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
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
      Left            =   11025
      TabIndex        =   30
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P.V.P"
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
      Left            =   9765
      TabIndex        =   28
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
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
      Left            =   8715
      TabIndex        =   26
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Left            =   12600
      TabIndex        =   32
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label LabelStock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999,999.99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   7350
      TabIndex        =   25
      Top             =   1995
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stock"
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
      Left            =   7350
      TabIndex        =   24
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NOMBRE DEL CLIENTE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1470
      TabIndex        =   2
      Top             =   105
      Width           =   10620
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CONVERSION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2940
      TabIndex        =   15
      Top             =   1260
      Width           =   1485
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COTIZACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label LabelStockArt 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " PRODUCTO"
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
      TabIndex        =   22
      Top             =   1680
      Width           =   7155
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " BODEGA"
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
      Left            =   8400
      TabIndex        =   20
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " LINEA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   16065
      TabIndex        =   8
      Top             =   105
      Width           =   2745
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " FECHA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "FacturasPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Grupo_Inv As String

Dim Total_PV As Currency
Dim SaldoPendiente As Currency

Dim CantSaldoPendiente As Integer

Dim ParpadearSaldo As Boolean
Dim SiTPCJ As Boolean
Dim SiTPBA As Boolean

Private Sub Command2_Click()
  FrmRifa.Visible = False
  TextCant.SetFocus
End Sub

Private Sub Command4_Click()
  TxtRifaD = "0"
  TxtRifaH = "0"
  FrmRifa.Visible = False
End Sub

Private Sub DCArticulo_GotFocus()
    Calculos_Totales_Factura FA
    LabelSubTotal.Caption = Format(FA.Sin_IVA, "#,##0.00")
    LabelConIVA.Caption = Format(FA.Con_IVA, "#,##0.00")
    LabelIVA.Caption = Format(FA.Total_IVA, "#,##0.00")
    LabelTotal.Caption = Format(FA.Total_MN, "#,##0.00")
End Sub

Private Sub DCArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    Keys_Especiales Shift
    Select Case KeyCode
      Case vbKeyEscape
           Calculos_Totales_Factura FA
           LabelSubTotal.Caption = Format(FA.Sin_IVA, "#,##0.00")
           LabelConIVA.Caption = Format(FA.Con_IVA, "#,##0.00")
           LabelIVA.Caption = Format(FA.Total_IVA, "#,##0.00")
           LabelTotal.Caption = Format(FA.Total_MN, "#,##0.00")
           If FA.TC = "DO" Then
              TxtEfectivo = "0.00"
              Total_PV = Val(InputBox("INGRESE EL TOTAL RECIBIDO", "INGRESO TOTAL", "0"))
              ReCalcular_PVP_Factura AdoAsientoF, Total_PV
           End If
           Calculos_Totales_Factura FA
           LabelSubTotal.Caption = Format(FA.Sin_IVA, "#,##0.00")
           LabelConIVA.Caption = Format(FA.Con_IVA, "#,##0.00")
           LabelIVA.Caption = Format(FA.Total_IVA, "#,##0.00")
           LabelTotal.Caption = Format(FA.Total_MN, "#,##0.00")
           TxtEfectivo.SetFocus
      Case vbKeyReturn
           SiguienteControl
      Case vbKeyF1
           With AdoArticulo.Recordset
            If .RecordCount Then
               .MoveFirst
               .Find ("Nom_Art = '" & DCArticulo & "' ")
                If Not .EOF Then MsgBox .fields("Producto") & ":" & vbCrLf & .fields("Ayuda")
            End If
           End With
    End Select
End Sub

Private Sub DCArticulo_LostFocus()
Dim Encontro As Boolean
    Encontro = Leer_Codigo_Inv(DCArticulo.Text, FechaSistema, Cod_Bodega)
    If Encontro Then
       DatosArticulos
''       If DatInv.Stock > 0 Then
'''          TextVUnit.Enabled = True
'''          TextCant.Enabled = True
''          TextCant.SetFocus
''       Else
''          TextCant.Enabled = False
''          TextVUnit.Enabled = False
''       End If
    Else
''       TextCant.Enabled = False
''       TextVUnit.Enabled = False
       DCArticulo.SetFocus
    End If
End Sub

Private Sub DCCliente_GotFocus()
   FA.DireccionS = Ninguno
End Sub

Private Sub DCCliente_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCCliente_KeyPress(KeyAscii As Integer)
Dim Busqueda As String
    Busqueda = DCCliente.Text
    If Len(Busqueda) > 1 Then
       sSQL = "SELECT TOP 50 Cliente,Codigo,CI_RUC,TD,Grupo,Direccion,DirNumero " _
            & "FROM Clientes "
       If IsNumeric(Busqueda) Then
          If Len(Busqueda) = 4 Then sSQL = sSQL & "WHERE DirNumero = '" & Busqueda & "' " Else sSQL = sSQL & "WHERE CI_RUC LIKE '" & Busqueda & "%' "
       Else
          sSQL = sSQL & "WHERE Cliente LIKE '%" & Busqueda & "%' "
       End If
       sSQL = sSQL & "ORDER BY Cliente "
       Select_Adodc AdoBenef, sSQL
    End If
End Sub

Private Sub DCCliente_LostFocus()
Dim DireccionC As String
  SaldoPendiente = 0
  CantSaldoPendiente = 0
  With AdoBenef.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       If IsNumeric(DCCliente.Text) Then
         .Find ("CI_RUC = '" & DCCliente.Text & "'")
       Else
         .Find ("Cliente = '" & DCCliente.Text & "'")
       End If
       If Not .EOF Then
          CodigoBenef = .fields("Codigo")
          CodigoCliente = .fields("Codigo")
          NombreCliente = .fields("Cliente")
          Grupo_No = .fields("Grupo")
          TipoDoc = .fields("TD")
          DireccionC = .fields("Direccion")
          DCCliente.Text = .fields("Cliente")
          LblRUC.Caption = .fields("CI_RUC")
                    
          sSQL = "SELECT Direccion " _
               & "FROM Clientes_Datos_Extras " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Codigo = '" & CodigoCliente & "' " _
               & "AND Tipo_Dato = 'DIRECCION' " _
               & "ORDER BY Direccion, Fecha_Registro DESC "
          SelectDB_Combo DCDireccion, AdoDireccion, sSQL, "Direccion", , "Clientes_Datos_Extras"
          If AdoDireccion.Recordset.RecordCount <= 0 Then DCDireccion = DireccionC
          
          sSQL = "SELECT COUNT(Factura) CantFact, SUM(Saldo_MN) As TSaldo_MN " _
               & "FROM Facturas " _
               & "WHERE Item = '" & NumEmpresa & "' " _
               & "AND Periodo = '" & Periodo_Contable & "' " _
               & "AND CodigoC = '" & CodigoCliente & "' "
          Select_Adodc AdoAux, sSQL
          If AdoAux.Recordset.RecordCount > 0 Then
             If Not IsNull(AdoAux.Recordset.fields("TSaldo_MN")) Then
                SaldoPendiente = AdoAux.Recordset.fields("TSaldo_MN")
                CantSaldoPendiente = AdoAux.Recordset.fields("CantFact")
             End If
          End If
       Else
          NombreCliente = DCCliente.Text
          FacturasPV.Visible = False
          Nuevo = True
          MsgBox "Cliente No existe"
          FClientesFlash.Show 1
          FacturasPV.Visible = True
          Listar_Clientes_PV
       End If
   Else
       NombreCliente = DCCliente.Text
       Nuevo = True
       FClientesFlash.Show 1
       Listar_Clientes_PV
   End If
  End With
End Sub

Private Sub DCDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub DCDireccion_LostFocus()
Dim DireccionAux As String
    DireccionAux = UCaseStrg(MidStrg(DCDireccion, 1, 80))
    sSQL = "SELECT " & Full_Fields("Clientes_Datos_Extras") & " " _
         & "FROM Clientes_Datos_Extras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Codigo = '" & CodigoCliente & "' " _
         & "AND Direccion = '" & DireccionAux & "' " _
         & "AND Tipo_Dato = 'DIRECCION' "
    Select_Adodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount <= 0 Then
         Titulo = "Formulario de Grabacion"
         Mensajes = "Esta Direccion no esta registrada: " & vbCrLf _
                  & DireccionAux & vbCrLf _
                  & "Desea registrarla?"
         If BoxMensaje = vbYes Then
            SetAddNew AdoAux
            SetFields AdoAux, "Tipo_Dato", "DIRECCION"
            SetFields AdoAux, "Codigo", CodigoCliente
            SetFields AdoAux, "Direccion", DireccionAux
            SetUpdate AdoAux
            DCDireccion = DireccionAux
         Else
            DCCliente.SetFocus
         End If
     End If
    End With
    FA.DireccionS = DCDireccion
End Sub

Private Sub MBFecha_GotFocus()
  MarcarTexto MBFecha
End Sub

Private Sub MBFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Public Sub Listar_Clientes_PV()
  sSQL = "SELECT TOP 50 Cliente,Codigo,CI_RUC,TD,Grupo,Direccion " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "AND FA <> " & Val(adFalse) & " " _
       & "UNION " _
       & "SELECT Cliente,Codigo,CI_RUC,TD,Grupo,Direccion " _
       & "FROM Clientes " _
       & "WHERE Codigo = '9999999999' " _
       & "ORDER BY Cliente "
  SelectDB_Combo DCCliente, AdoBenef, sSQL, "Cliente"
End Sub

Private Sub Command1_Click()
  Unload FacturasPV
End Sub

Private Sub Command3_Click()
  FechaValida MBFecha
  FechaTexto = MBFecha
  Calculos_Totales_Factura FA
  If (FA.Total_MN - Val(CCur(TxtEfectivo))) >= 0 Then
     Mensajes = "Esta Seguro que desea grabar: " & vbCrLf _
              & "Comprobante (" & FA.TC & ")  No. " & TextFacturaNo.Text
     Titulo = "Formulario de Grabacion"
     If BoxMensaje = vbYes Then
        FA.Nota = TxtNota
        FA.Observacion = TxtObservacion
        Moneda_US = False
        TextoFormaPago = PagoCont
        ProcGrabar
        
        LblSerie.Caption = SerieFactura & "-"
        NumComp = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
        TextFacturaNo.Text = Format$(NumComp, "000000000")

        sSQL = "DELETE * " _
             & "FROM Asiento_F " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        Ejecutar_SQL_SP sSQL
        
        sSQL = "SELECT * " _
             & "FROM Asiento_F " _
             & "WHERE Item = '" & NumEmpresa & "' " _
             & "AND CodigoU = '" & CodigoUsuario & "' "
        Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL
        Ln_No = 1
        DCArticulo.SetFocus
     End If
  Else
     MsgBox "Error: El Efectivo no alcanza para grabar"
  End If
End Sub

Private Sub DCBodega_LostFocus()
  Cod_Bodega = Ninguno
  With AdoBodega.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Bodega = '" & DCBodega.Text & "' ")
       If Not .EOF Then Cod_Bodega = .fields("CodBod")
   End If
  End With
End Sub

Private Sub DGAsientoF_BeforeDelete(Cancel As Integer)
  Mensajes = "¿Realmente desea eliminar el campo " & vbCrLf & "(" _
           & AdoAsientoF.Recordset.fields("CODIGO") & ") " _
           & AdoAsientoF.Recordset.fields("PRODUCTO") & "   TOTAL -> " _
           & AdoAsientoF.Recordset.fields("TOTAL") & "?"
  Titulo = "Confirmación de eliminación"
  If BoxMensaje = 6 Then Cancel = False Else Cancel = True
End Sub

Private Sub Form_Activate()
  Grupo_Inv = Ninguno
  Ln_No = 1
  OpcMult.value = True
  Cant_Item_PV = 50
  sSQL = "DELETE * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  Ejecutar_SQL_SP sSQL
  CodigoCliente = "9999999999"
  NombreCliente = "CONSUMIDOR FINAL"
  DireccionCli = " S/N"
  DCCliente.Text = NombreCliente
  FacturasPV.Caption = FacturasPV.Caption & " (" & TipoFactura & ")"
  Label23.Caption = " Total Tarifa " & Porc_IVA * 100 & "%"
  Label3.Caption = " Total I.V.A. " & Porc_IVA * 100 & "%"
  TextCant.Text = "0"
  TextVUnit.Text = "0"
  LabelVTotal.Caption = "0"
  Modificar = False
  Bandera = True
  Mifecha = BuscarFecha(FechaSistema)
  SerieFactura = Leer_Campo_Empresa("Serie_FA")
  If Len(SerieFactura) < 6 Then SerieFactura = "999999"
  FA.TC = TipoFactura
  FA.Serie = SerieFactura
  CodigoL = Ninguno
  Cta_Cobrar = Ninguno
  Autorizacion = "9999999999"
  sSQL = "SELECT CxC, Codigo, Autorizacion, Fact, Serie " _
       & "FROM Catalogo_Lineas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Fact = '" & TipoFactura & "' " _
       & "" _
       & "AND TL <> " & Val(adFalse) & " " _
       & "ORDER BY Codigo "
  Select_Adodc AdoLinea, sSQL
  With AdoLinea.Recordset
    If .RecordCount > 0 Then
        Cta_Cobrar = .fields("CxC")
        CodigoL = .fields("Codigo")
        Autorizacion = .fields("Autorizacion")
        FA.TC = .fields("Fact")
        FA.Serie = .fields("Serie")
        FA.Cta_CxP = Cta_Cobrar
        FA.Cod_CxC = CodigoL
        FA.Autorizacion = Autorizacion
    Else
        MsgBox "Falta Organizar la CxC en Puntos de Venta." & vbCrLf _
             & "Salga de este proceso y llame al su técnico" & vbCrLf _
             & "o al Contador de su Organizacion."
    End If
  End With
  
  Select Case TipoFactura
    Case "PV"
        FacturasPV.Caption = "INGRESAR TICKET"
        Label1.Caption = " TICKET No."
        Label3.Caption = " I.V.A. " & Format$(Porc_IVA * 100, "#0.00") & "%"
    Case "CP"
        FacturasPV.Caption = "INGRESAR CHEQUES PROTESTADOS"
        Label1.Caption = " COMPROBANTE No."
        Label3.Caption = " I.V.A. 0.00%"
    Case "NV"
        FacturasPV.Caption = "INGRESAR NOTA DE VENTA"
        Label1.Caption = " NOTA DE VENTA No."
        Label3.Caption = " I.V.A. 0.00%"
    Case "DO"
        FacturasPV.Caption = "INGRESAR NOTA DE DONACION"
        Label1.Caption = " NOTA DE DONACION No."
        Label3.Caption = " I.V.A. 0.00%"
    Case "LC"
        FacturasPV.Caption = "INGRESAR LIQUIDACION DE COMPRAS"
        Label1.Caption = " LIQUIDACION DE COMPRAS No."
        Label3.Caption = " I.V.A. 0.00%"
        OpcDiv.value = True
        'If Len(Opc_Grupo_Div) > 1 Then Grupo_Inv = Opc_Grupo_Div
    Case Else
        FacturasPV.Caption = "INGRESAR FACTURA"
        Label1.Caption = " FACTURA No."
        Label3.Caption = " I.V.A. " & Format$(Porc_IVA * 100, "#0.00") & "%"
  End Select
  
  LblSerie.Caption = SerieFactura & "-"
  NumComp = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, False)
  TextFacturaNo.Text = Format$(NumComp, "000000000")
  
  sSQL = "SELECT * " _
       & "FROM Asiento_F " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND CodigoU = '" & CodigoUsuario & "' "
  SQLDec = "PRECIO 4|CORTE 5|."
  Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL, SQLDec
  RatonNormal
  Listar_Clientes_PV
  MBFecha.Text = FechaSistema
  Cod_Bodega = Ninguno
  sSQL = "SELECT CodBod, Bodega, ID " _
       & "FROM Catalogo_Bodegas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "ORDER BY CodBod "
  SelectDB_Combo DCBodega, AdoBodega, sSQL, "Bodega"
  With AdoBodega.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
      .Find ("Bodega = '" & DCBodega.Text & "' ")
       If Not .EOF Then Cod_Bodega = .fields("CodBod")
   End If
  End With
  
  sSQL = "SELECT Codigo, DC " _
       & "FROM Ctas_Proceso " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND Detalle = 'Cta_Caja_GMN' "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     Cta_CajaG = AdoAux.Recordset.fields("Codigo")
     DC_CajaG = AdoAux.Recordset.fields("DC")
  End If

  SiTPCJ = False
  sSQL = "SELECT Codigo " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'CJ' " _
       & "AND TP = '" & TipoFactura & "' " _
       & "ORDER BY Codigo "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     Cta_CajaG = AdoAux.Recordset.fields("Codigo")
     SiTPCJ = True
  End If
  
  SiTPBA = False
  sSQL = "SELECT TOP 1 Codigo " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'BA' " _
       & "AND TP = '" & TipoFactura & "' "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then SiTPBA = True
  
  sSQL = "SELECT Codigo & Space(2) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND DG = 'D' " _
       & "AND TC = 'CJ' "
  If SiTPCJ Then sSQL = sSQL & "AND TP = '" & TipoFactura & "' "
  sSQL = sSQL _
       & "UNION " _
       & "SELECT Codigo & Space(2) & Cuenta As NomCuenta " _
       & "FROM Catalogo_Cuentas " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND DG = 'D' "
  If SiTPBA Then
     sSQL = sSQL _
          & "AND TC = 'BA' " _
          & "AND TP = '" & TipoFactura & "' "
  Else
     sSQL = sSQL & "AND TC IN ('BA','C','P') "
  End If
  sSQL = sSQL & "ORDER BY NomCuenta "
  SelectDB_Combo DCBanco, AdoBanco, sSQL, "NomCuenta"
  
  sSQL = "SELECT Producto & ' -> ' & Codigo_Inv As Nom_Art,* " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'I' "
  If Len(Grupo_Inv) > 1 Then sSQL = sSQL & "AND MidStrg(Codigo_Inv,1,2) = '" & Grupo_Inv & "' "
  sSQL = sSQL & "ORDER BY Producto,Codigo_Inv "
  SelectDB_List DLGrupo, AdoGrupo, sSQL, "Nom_Art"
  
  TextCotiza.Text = Format(Dolar, "#,##0.00")
  TextCotiza.Enabled = False
  OpcMult.Enabled = False
  OpcDiv.Enabled = False
  sSQL = "SELECT TOP 1 Div " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' " _
       & "AND Div <> 0 "
  Select_Adodc AdoAux, sSQL
  If AdoAux.Recordset.RecordCount > 0 Then
     TextCotiza.Enabled = True
     OpcMult.Enabled = True
     OpcDiv.Enabled = True
  End If
  
  sSQL = "SELECT Producto,Codigo_Inv,Codigo_Barra " _
       & "FROM Catalogo_Productos " _
       & "WHERE Item = '" & NumEmpresa & "' " _
       & "AND Periodo = '" & Periodo_Contable & "' " _
       & "AND TC = 'P' "
  If Len(Grupo_Inv) > 1 Then sSQL = sSQL & "AND MidStrg(Codigo_Inv,1,2) = '" & Grupo_Inv & "' "
  If TipoFactura = "CP" Then
     sSQL = sSQL & "AND Cta_Inventario = '0' "
  Else
     sSQL = sSQL & "AND LEN(Cta_Inventario) > 1 "
  End If
  sSQL = sSQL & "ORDER BY Producto,Codigo_Inv "
  SelectDB_Combo DCArticulo, AdoArticulo, sSQL, "Producto"
  If TipoFactura <> "CP" Then
     ReDim ExisteCtas(4) As String
     ExisteCtas(0) = Cta_Cobrar
     ExisteCtas(1) = Cta_CajaG
     ExisteCtas(2) = Cta_CajaGE
     ExisteCtas(3) = Cta_CajaBA
     VerSiExisteCta ExisteCtas
  End If
  Timer1.Interval = 1000
  SaldoPendiente = 0
  CantSaldoPendiente = 0
  ParpadearSaldo = True
  FacturasPV.WindowState = 2
  If AdoArticulo.Recordset.RecordCount <= 0 Then
     MsgBox "No existen Productos de Venta"
     Unload FacturasPV
  End If
End Sub

Private Sub Form_Deactivate()
  FacturasPV.WindowState = 1
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoBanco
  ConectarAdodc AdoBenef
  ConectarAdodc AdoGrupo
  ConectarAdodc AdoLinea
  ConectarAdodc AdoBodega
  ConectarAdodc AdoFactura
  ConectarAdodc AdoArticulo
  ConectarAdodc AdoAsientoF
  ConectarAdodc AdoDireccion
  
  SRI_Obtener_Datos_Comprobantes_Electronicos
  
  Encerar_Factura FA
End Sub

Private Sub MBFecha_LostFocus()
  FechaValida MBFecha
  Validar_Porc_IVA MBFecha
  FechaTexto1 = MBFecha
End Sub

Private Sub TextCant_GotFocus()
  MarcarTexto TextCant
  Codigos = Ninguno
  
  If DatInv.Stock <= 0 Then
     Mensajes = "EL PRODUCTO:" & vbCrLf & vbCrLf & DCArticulo & vbCrLf & vbCrLf & "ES UN PRODUCTO SIN EXISTENCIA"
     Titulo = "PUNTO DE VENTA"
     MsgBox Mensajes, vbInformation, Titulo
     DCArticulo.SetFocus
  End If
End Sub

Private Sub TextCant_KeyDown(KeyCode As Integer, Shift As Integer)
  Keys_Especiales Shift
End Sub

Private Sub TextCant_LostFocus()
Dim DifStock As Single
  Cantidad = Val(TextCant)
  DifStock = Val(CCur(LabelStock.Caption)) - Cantidad
  If Redondear(DifStock, 2) < 0 Then
     Mensajes = UCaseStrg(DCArticulo) & vbCrLf & vbCrLf & "NO PUEDE QUEDAR EXISTENCIA NEGATIVA, SOLICITE ALIMENTACION DE STOCK"
     Titulo = "PUNTO DE VENTA"
     MsgBox Mensajes, vbInformation, "PUNTO DE VENTA"
     DCArticulo.SetFocus
  End If
End Sub

Private Sub TextCotiza_GotFocus()
  TextoValido TextCotiza
End Sub

Private Sub TextCotiza_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_Change()
   If IsNumeric(TextVUnit) And IsNumeric(TextCant) Then
      If Val(TextVUnit) = 0 Then TextVUnit = "0.01"
      If OpcMult.value Then Real1 = CCur(TextCant) * CCur(TextVUnit) Else Real1 = CCur(TextCant) / CCur(TextVUnit)
      LabelVTotal.Caption = Format$(Real1, "#,##0.0000")
   Else
      LabelVTotal.Caption = "0.0000"
   End If
End Sub

Private Sub TextVUnit_GotFocus()
  MarcarTexto TextVUnit
End Sub

Private Sub TextVUnit_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Total_PV As Currency
   Keys_Especiales Shift
   If CtrlDown And KeyCode = vbKeyP Then
      Total_PV = InputBox("INGRESE PRECIO TOTAL", "INGRESO TOTAL", TextVUnit)
      If Val(TextCant) = 0 Then TextCant = "1"
      TextVUnit = Redondear(Total_PV / Val(TextCant), 8)
   End If
   PresionoEnter KeyCode
End Sub

Private Sub TextVUnit_LostFocus()
   If TipoFactura = "CP" Then TextoValido TextVUnit, True, , 2 Else TextoValido TextVUnit, True, , 4
   If Val(TextVUnit) <= 0 Then TextVUnit = Format$(Precio, "#,##0.0000")
End Sub

Public Sub ProcGrabar()
 DGAsientoF.Visible = False
 FA.Porc_IVA = Porc_IVA
 FA.Gavetas = Val(TxtGavetas)
 'Seteamos los encabezados para las facturas
  Calculos_Totales_Factura FA
  If AdoAsientoF.Recordset.RecordCount > 0 Then
     RatonReloj
     FechaTexto = MBFecha
     FA.Fecha = FechaTexto
     FA.CodigoC = CodigoCliente
     HoraTexto = Format$(Time, FormatoTimes)
     Total_FacturaME = 0
     Moneda_US = False
     Total_Factura = Redondear(FA.Sin_IVA + FA.Con_IVA - FA.Descuento - FA.Descuento2 + FA.Total_IVA + FA.Servicio, 2)
     Total_FacturaME = Total_Factura
     If Moneda_US Then Total_Factura = Redondear(Total_Factura * Dolar, 2) Else Total_FacturaME = 0
     Saldo = Total_Factura
     Saldo_ME = Total_FacturaME
     If Saldo < 0 Then Saldo = 0
     FA.Nuevo_Doc = True
     Factura_No = ReadSetDataNum(FA.TC & "_SERIE_" & FA.Serie, True, True)
     FA.Factura = Factura_No
     TipoFactura = FA.TC
     Select Case TipoFactura
       Case "PV": Control_Procesos "F", "Grabar Ticket No. " & Factura_No
       Case "NV": Control_Procesos "F", "Grabar Nota de Venta No. " & Factura_No
       Case "CP": Control_Procesos "F", "Grabar Cheque Protestado No. " & Factura_No
       Case "LC": Control_Procesos "F", "Grabar Liquidacion de Compras No. " & Factura_No
       Case "DO": Control_Procesos "F", "Grabar Nota de Donacion No. " & Factura_No
       Case Else: Control_Procesos "F", "Grabar Factura No. " & Factura_No
     End Select
     
     sSQL = "DELETE * " _
          & "FROM Detalle_Factura " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TipoFactura & "' "
     Ejecutar_SQL_SP sSQL
     sSQL = "DELETE * " _
          & "FROM Facturas " _
          & "WHERE Factura = " & Factura_No & " " _
          & "AND Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND TC = '" & TipoFactura & "' "
     Ejecutar_SQL_SP sSQL
     TextoFormaPago = PagoCred
     T = Pendiente
    'Grabamos el numero de factura
     RatonNormal
     Grabar_Factura FA, True
     
     If FA.TC <> "CP" Then
        Evaluar = True
        FechaTexto = MBFecha.Text
        
       'Llenamos datos genericos de los Abonos de la  factura a grabar
        DiarioCaja = ReadSetDataNum("Recibo_No", True, True)
        TA.T = Normal
        TA.Fecha = FA.Fecha
        TA.Cta_CxP = FA.Cta_CxP
        TA.TP = FA.TC
        TA.Serie = FA.Serie
        TA.Factura = FA.Factura
        TA.Autorizacion = FA.Autorizacion
        TA.CodigoC = FA.CodigoC
        
       'Abono en efectivo
        TA.Cta = Cta_CajaG
        TA.Banco = "EFECTIVO MN"
        TA.Cheque = Format$(FA.Factura, "00000000")
        TA.Abono = Val(TxtEfectivo.Text)
        Grabar_Abonos TA
        
       'Abono de Factura Banco
        TA.Cta = TrimStrg(SinEspaciosIzq(DCBanco))
        TA.Banco = TextBanco
        TA.Cheque = TextCheqNo
        TA.Abono = Val(TextCheque.Text)
        Grabar_Abonos TA
     End If
             
    'Actualizamos el saldo de la factura
     Actualizar_Saldos_Facturas_SP FA.TC, FA.Serie, FA.Factura
    'Mayorizar_Inventario_SP
     
     sSQL = "DELETE * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     Ejecutar_SQL_SP sSQL
     
     sSQL = "SELECT * " _
          & "FROM Asiento_F " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND CodigoU = '" & CodigoUsuario & "' "
     Select_Adodc_Grid DGAsientoF, AdoAsientoF, sSQL
     Select Case FA.TC
       Case "FA", "NV", "DO"
            If Len(FA.Autorizacion) >= 13 Then
               If FA.TC = "DO" Then
                  Generar_PDF_Donacion FA, True
               ElseIf FA.TC = "NV" Then
                  Imprimir_Punto_Venta FA
               Else
                  SRI_Crear_Clave_Acceso_Facturas FA, True
               End If
            Else
               Imprimir_Punto_Venta FA
''               FA.Desde = FA.Factura
''               FA.Hasta = FA.Factura
''               Imprimir_Facturas_CxC FacturasPV, FA, True, False, True, True
            End If
     End Select
     Listar_Clientes_PV
     NombreCliente = "CONSUMIDOR FINAL"
     DCCliente.Text = NombreCliente
  Else
     MsgBox "No se puede grabar la Factura," & vbCrLf & "falta datos."
  End If
  DGAsientoF.Visible = True
End Sub

Public Sub DatosArticulos()
''  With AdoArticulo.Recordset
''   If .RecordCount > 0 Then
    Codigos = DatInv.Codigo_Inv
    Producto = DatInv.Producto
    Cta_Ventas = DatInv.Cta_Ventas
    Precio = DatInv.PVP
    LabelStock.Caption = Format$(DatInv.Stock, "#,##0.00")
    BanIVA = DatInv.IVA
    TextVUnit = Format$(Precio, "#,##0.0000")
    If TipoFactura = "NV" Then BanIVA = False
    LabelStockArt.Caption = "PRODUCTO" & String(93 - Len(DatInv.Codigo_Inv), " ") & DatInv.Codigo_Inv
    'DCArticulo.Text = Producto
''   End If
''  End With
End Sub

Private Sub Timer1_Timer()
  If ParpadearSaldo And SaldoPendiente <> 0 Then
     LblSaldoPendiente.Caption = Format(SaldoPendiente, "#,##0.00")
     Label6.Caption = "NOMBRE DEL CLIENTE:" & String(70, " ") & "CANTIDAD DOCUMENTOS PENDIENTES: " & CantSaldoPendiente
  Else
     LblSaldoPendiente.Caption = ""
     Label6.Caption = "NOMBRE DEL CLIENTE:"
  End If
  ParpadearSaldo = Not ParpadearSaldo
End Sub

Private Sub TxtDocumentos_GotFocus()
  MarcarTexto TxtDocumentos
End Sub

Private Sub TxtDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
   Keys_Especiales Shift
   PresionoEnter KeyCode
   If CtrlDown And KeyCode = vbKeyR Then
      FrmRifa.Visible = True
      TxtRifaD.SetFocus
   End If
End Sub

Private Sub TxtDocumentos_LostFocus()
Dim Grabar_PV As Boolean
Dim ProductoAux As String
  
   TextoValido TxtDocumentos
   
   Grabar_PV = True
   If Cant_Item_PV > 0 And (AdoAsientoF.Recordset.RecordCount > Cant_Item_PV) Then Grabar_PV = False
   
   'MsgBox Cant_Item_PV
   If Grabar_PV Then
      LabelVTotal.Caption = Format$(Real1, "#,##0.0000")
      Real1 = 0: Real2 = 0: Real3 = 0
      If IsNumeric(TextVUnit) And IsNumeric(TextCant) Then
         'If Val(TextVUnit) = 0 Then TextVUnit = "0.01"
         'If Val(TextCant) = 0 Then TextCant = "1"
         If OpcMult.value Then
            Real1 = CCur(TextCant) * CCur(TextVUnit)
         Else
            If CCur(TextVUnit) <> 0 Then Real1 = CCur(TextCant) / CCur(TextVUnit)
         End If
      End If
      If Real1 > 0 Then
         Select Case TipoFactura
           Case "NV", "PV": Real3 = 0
           Case Else
                If BanIVA Then Real3 = Redondear((Real1 - Real2) * Porc_IVA, 2) Else Real3 = 0
         End Select
         LabelVTotal.Caption = Format$(Real1, "#,##0.00")
         If Len(TxtDocumentos) > 1 Then Producto = Producto & " - " & TxtDocumentos
         If IsNumeric(TxtRifaD) And IsNumeric(TxtRifaH) And Val(TxtRifaD) < Val(TxtRifaH) Then
               For I = Val(TxtRifaD) To Val(TxtRifaH)
                   ProductoAux = Producto & " " & Format(I, "000000")
                   SetAddNew AdoAsientoF
                   SetFields AdoAsientoF, "CODIGO", DatInv.Codigo_Inv
                   SetFields AdoAsientoF, "CODIGO_L", CodigoL
                   SetFields AdoAsientoF, "PRODUCTO", MidStrg(ProductoAux, 1, 150)
                   SetFields AdoAsientoF, "Tipo_Hab", MidStrg(TxtDocumentos, 1, 12)
                   SetFields AdoAsientoF, "CANT", 1
                   SetFields AdoAsientoF, "PRECIO", CCur(TextVUnit)
                   SetFields AdoAsientoF, "TOTAL", Real1
                   SetFields AdoAsientoF, "Total_IVA", Real3
                   SetFields AdoAsientoF, "Item", NumEmpresa
                   SetFields AdoAsientoF, "CodigoU", CodigoUsuario
                   SetFields AdoAsientoF, "CodBod", Cod_Bodega
                   SetFields AdoAsientoF, "A_No", Ln_No
                   SetUpdate AdoAsientoF
                   Ln_No = Ln_No + 1
               Next I
         Else
            If CCur(TextCant) > 0 And DatInv.Stock > 0 Then
               SetAddNew AdoAsientoF
               SetFields AdoAsientoF, "CODIGO", DatInv.Codigo_Inv
               SetFields AdoAsientoF, "CODIGO_L", CodigoL
               SetFields AdoAsientoF, "PRODUCTO", MidStrg(Producto, 1, 150)
               SetFields AdoAsientoF, "Tipo_Hab", MidStrg(TxtDocumentos, 1, 12)
               SetFields AdoAsientoF, "CANT", CCur(TextCant)
               SetFields AdoAsientoF, "PRECIO", CCur(TextVUnit)
               SetFields AdoAsientoF, "TOTAL", Real1
               SetFields AdoAsientoF, "Total_IVA", Real3
               SetFields AdoAsientoF, "Item", NumEmpresa
               SetFields AdoAsientoF, "CodigoU", CodigoUsuario
               SetFields AdoAsientoF, "CodBod", Cod_Bodega
               SetFields AdoAsientoF, "A_No", Ln_No
               SetFields AdoAsientoF, "COSTO", DatInv.Costo
               If DatInv.Costo > 0 Then
                  SetFields AdoAsientoF, "Cta_Inv", DatInv.Cta_Inventario
                  SetFields AdoAsientoF, "Cta_Costo", DatInv.Cta_Costo_Venta
               End If
               SetUpdate AdoAsientoF
               Ln_No = Ln_No + 1
            End If
         End If
      End If
   Else
      MsgBox "Ya no puede ingresar mas productos"
      'TxtEfectivo.SetFocus
   End If
   TextCant.Text = "0"
   DCArticulo.SetFocus
End Sub

Private Sub TxtEfectivo_GotFocus()
    TxtEfectivo = Format(FA.Total_MN, "#,##0.00")
    MarcarTexto TxtEfectivo
    LabelSubTotal.Caption = Format(FA.Sin_IVA, "#,##0.00")
    LabelConIVA.Caption = Format(FA.Con_IVA, "#,##0.00")
    LabelIVA.Caption = Format(FA.Total_IVA, "#,##0.00")
    LabelTotal.Caption = Format(FA.Total_MN, "#,##0.00")
End Sub

Private Sub TxtEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtEfectivo_Change()
  If Val(TextCotiza) > 0 Then
     If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format$(Val(TxtEfectivo) - Total_FacturaME, "#,##0.00")
  Else
     If Val(TxtEfectivo) > 0 Then LblCambio.Caption = Format$(Val(TxtEfectivo) - Total_Factura, "#,##0.00")
  End If
End Sub

Private Sub TxtEfectivo_LostFocus()
  TextoValido TxtEfectivo, True, , 2
  If Val(TextCotiza) > 0 Then
     LblCambio.Caption = Format$(Val(CCur(TxtEfectivo)) - Total_FacturaME, "#,##0.00")
     If (Val(CCur(TxtEfectivo)) - Total_FacturaME) >= 0 Then Command3.SetFocus
  Else
     LblCambio.Caption = Format$(Val(CCur(TxtEfectivo)) - Total_Factura, "#,##0.00")
     If (Val(CCur(TxtEfectivo)) - Total_Factura) >= 0 Then Command3.SetFocus
  End If
End Sub

Private Sub TxtGavetas_GotFocus()
  MarcarTexto TxtGavetas
End Sub

Private Sub TxtGavetas_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtGavetas_LostFocus()
  TextoValido TxtGavetas, True, , 0
End Sub

Private Sub TxtNota_GotFocus()
  MarcarTexto TxtNota
End Sub

Private Sub TxtNota_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtNota_LostFocus()
  TextoValido TxtNota, , True
End Sub

Private Sub TxtObservacion_GotFocus()
   MarcarTexto TxtObservacion
End Sub

Private Sub TxtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtObservacion_LostFocus()
  TextoValido TxtObservacion, , True
End Sub

Private Sub TxtRifaD_GotFocus()
  MarcarTexto TxtRifaD
End Sub

Private Sub TxtRifaD_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRifaD_LostFocus()
  TextoValido TxtRifaD, True, , 0
End Sub

Private Sub TxtRifaH_GotFocus()
  MarcarTexto TxtRifaH
End Sub

Private Sub TxtRifaH_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtRifaH_LostFocus()
  TextoValido TxtRifaH, True, , 0
End Sub
