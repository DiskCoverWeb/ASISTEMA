VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FModifica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificación"
   ClientHeight    =   5796
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   9192
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   9192
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTabModif 
      Height          =   5196
      Left            =   108
      TabIndex        =   0
      Top             =   108
      Width           =   8976
      _ExtentX        =   15833
      _ExtentY        =   9165
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   420
      BackColor       =   8421504
      TabCaption(0)   =   "Comprobante de Compra"
      TabPicture(0)   =   "FModifica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lbl(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lbl(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbl(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Lbl(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Lbl(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Lbl(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Lbl(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Lbl(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Lbl(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Lbl(13)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Lbl(14)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Lbl(15)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Lbl(16)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Lbl(17)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Lbl(18)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Lbl(38)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Lbl(39)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Lbl(40)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Lbl(41)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "MB(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "MB(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "MB(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "DC(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "DC(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "DC(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Txt(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Cmb(1)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Cmb(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Txt(0)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Txt(2)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Txt(3)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Txt(4)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Txt(5)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Txt(6)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Txt(7)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Txt(8)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Txt(9)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Cmb(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Txt(10)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Txt(11)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Txt(12)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Cmb(3)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "Concepto AIR"
      TabPicture(1)   =   "FModifica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Txt(18)"
      Tab(1).Control(1)=   "Txt(19)"
      Tab(1).Control(2)=   "Txt(20)"
      Tab(1).Control(3)=   "Txt(21)"
      Tab(1).Control(4)=   "Txt(22)"
      Tab(1).Control(5)=   "Txt(23)"
      Tab(1).Control(6)=   "Txt(24)"
      Tab(1).Control(7)=   "Txt(25)"
      Tab(1).Control(8)=   "Txt(26)"
      Tab(1).Control(9)=   "Txt(27)"
      Tab(1).Control(10)=   "Txt(28)"
      Tab(1).Control(11)=   "Txt(29)"
      Tab(1).Control(12)=   "Txt(30)"
      Tab(1).Control(13)=   "Txt(31)"
      Tab(1).Control(14)=   "Txt(32)"
      Tab(1).Control(15)=   "Txt(33)"
      Tab(1).Control(16)=   "Txt(34)"
      Tab(1).Control(17)=   "Txt(35)"
      Tab(1).Control(18)=   "Txt(36)"
      Tab(1).Control(19)=   "Txt(37)"
      Tab(1).Control(20)=   "Lbl(19)"
      Tab(1).Control(21)=   "Lbl(20)"
      Tab(1).Control(22)=   "Lbl(21)"
      Tab(1).Control(23)=   "Lbl(22)"
      Tab(1).Control(24)=   "Lbl(23)"
      Tab(1).Control(25)=   "Lbl(24)"
      Tab(1).Control(26)=   "Lbl(25)"
      Tab(1).Control(27)=   "Lbl(26)"
      Tab(1).Control(28)=   "Lbl(27)"
      Tab(1).Control(29)=   "Lbl(28)"
      Tab(1).Control(30)=   "Lbl(29)"
      Tab(1).Control(31)=   "Lbl(30)"
      Tab(1).Control(32)=   "Lbl(31)"
      Tab(1).Control(33)=   "Lbl(32)"
      Tab(1).Control(34)=   "Lbl(33)"
      Tab(1).Control(35)=   "Lbl(34)"
      Tab(1).Control(36)=   "Lbl(35)"
      Tab(1).Control(37)=   "Lbl(36)"
      Tab(1).Control(38)=   "Lbl(37)"
      Tab(1).ControlCount=   39
      Begin VB.ComboBox Cmb 
         Height          =   288
         Index           =   3
         Left            =   3780
         TabIndex        =   85
         Text            =   "Combo1"
         Top             =   3672
         Width           =   876
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   336
         Index           =   12
         Left            =   5940
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   3564
         Width           =   1092
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   336
         Index           =   11
         Left            =   1620
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   3672
         Width           =   1092
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   336
         Index           =   10
         Left            =   5940
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   3132
         Width           =   1092
      End
      Begin VB.ComboBox Cmb 
         Height          =   288
         Index           =   2
         Left            =   3780
         TabIndex        =   81
         Text            =   "Combo1"
         Top             =   3240
         Width           =   876
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   336
         Index           =   9
         Left            =   1620
         TabIndex        =   80
         Text            =   "Text1"
         Top             =   3240
         Width           =   1092
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   336
         Index           =   8
         Left            =   3780
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   2808
         Width           =   1092
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   336
         Index           =   7
         Left            =   1620
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   2808
         Width           =   1092
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   336
         Index           =   6
         Left            =   5940
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   2376
         Width           =   1092
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   336
         Index           =   5
         Left            =   3780
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   2376
         Width           =   1092
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   336
         Index           =   4
         Left            =   1620
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   2376
         Width           =   1092
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   3
         Left            =   5184
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   1512
         Width           =   660
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   2
         Left            =   3240
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   1512
         Width           =   660
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   0
         Left            =   7560
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   540
         Width           =   660
      End
      Begin VB.ComboBox Cmb 
         Height          =   288
         Index           =   0
         Left            =   6480
         TabIndex        =   64
         Text            =   "Combo1"
         Top             =   540
         Width           =   876
      End
      Begin VB.ComboBox Cmb 
         Height          =   288
         Index           =   1
         Left            =   7452
         TabIndex        =   55
         Text            =   "Combo1"
         Top             =   972
         Width           =   876
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   18
         Left            =   -70140
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   1620
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   19
         Left            =   -70140
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   1944
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   20
         Left            =   -70140
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   2268
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   21
         Left            =   -70140
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   2592
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   22
         Left            =   -70140
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   2916
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   23
         Left            =   -70140
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   3240
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   24
         Left            =   -70140
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   3564
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   25
         Left            =   -70140
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   3888
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   26
         Left            =   -70140
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   4212
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   27
         Left            =   -70140
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   4536
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   28
         Left            =   -73812
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   648
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   29
         Left            =   -73812
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   972
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   30
         Left            =   -73812
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1296
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   31
         Left            =   -73812
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1620
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   32
         Left            =   -73812
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1944
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   33
         Left            =   -73812
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2268
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   34
         Left            =   -73812
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2592
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   35
         Left            =   -73812
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2916
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   36
         Left            =   -73812
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   3240
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   37
         Left            =   -73812
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   3564
         Width           =   984
      End
      Begin VB.TextBox Txt 
         Height          =   336
         Index           =   1
         Left            =   1620
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1512
         Width           =   660
      End
      Begin MSDataListLib.DataCombo DC 
         Bindings        =   "FModifica.frx":0038
         DataSource      =   "AdoAux"
         Height          =   288
         Index           =   1
         Left            =   1620
         TabIndex        =   61
         ToolTipText     =   $"FModifica.frx":004D
         Top             =   864
         Width           =   4272
         _ExtentX        =   7535
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DC 
         Bindings        =   "FModifica.frx":00F5
         DataSource      =   "AdoClientes"
         Height          =   288
         Index           =   0
         Left            =   1620
         TabIndex        =   62
         ToolTipText     =   $"FModifica.frx":010F
         Top             =   540
         Width           =   4272
         _ExtentX        =   7535
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DC 
         Bindings        =   "FModifica.frx":01B7
         DataSource      =   "AdoTipoComp"
         Height          =   288
         Index           =   2
         Left            =   1620
         TabIndex        =   63
         ToolTipText     =   $"FModifica.frx":01D8
         Top             =   1188
         Width           =   4272
         _ExtentX        =   7535
         _ExtentY        =   508
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MB 
         Height          =   336
         Index           =   0
         Left            =   1620
         TabIndex        =   69
         ToolTipText     =   $"FModifica.frx":0280
         Top             =   1944
         Width           =   1176
         _ExtentX        =   2074
         _ExtentY        =   593
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "0"
      End
      Begin MSMask.MaskEdBox MB 
         Height          =   336
         Index           =   1
         Left            =   3780
         TabIndex        =   70
         ToolTipText     =   $"FModifica.frx":0308
         Top             =   1944
         Width           =   1176
         _ExtentX        =   2074
         _ExtentY        =   593
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "0"
      End
      Begin MSMask.MaskEdBox MB 
         Height          =   336
         Index           =   2
         Left            =   5940
         TabIndex        =   71
         ToolTipText     =   $"FModifica.frx":0390
         Top             =   1944
         Width           =   1176
         _ExtentX        =   2074
         _ExtentY        =   593
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "0"
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   41
         Left            =   5076
         TabIndex        =   79
         Top             =   3672
         Width           =   1308
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   40
         Left            =   2808
         TabIndex        =   78
         Top             =   3672
         Width           =   1308
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   39
         Left            =   216
         TabIndex        =   77
         Top             =   3780
         Width           =   1308
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   38
         Left            =   5076
         TabIndex        =   76
         Top             =   3240
         Width           =   1308
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   18
         Left            =   2808
         TabIndex        =   60
         Top             =   3240
         Width           =   1308
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   17
         Left            =   216
         TabIndex        =   59
         Top             =   3240
         Width           =   660
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   16
         Left            =   2808
         TabIndex        =   58
         Top             =   2808
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   15
         Left            =   216
         TabIndex        =   57
         Top             =   2808
         Width           =   1524
      End
      Begin VB.Label Lbl 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         Height          =   228
         Index           =   14
         Left            =   5076
         TabIndex        =   56
         Top             =   2484
         Width           =   660
      End
      Begin VB.Label Lbl 
         Caption         =   "Label20"
         Height          =   228
         Index           =   19
         Left            =   -71004
         TabIndex        =   54
         Top             =   2052
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label21"
         Height          =   228
         Index           =   20
         Left            =   -71004
         TabIndex        =   53
         Top             =   2376
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label22"
         Height          =   228
         Index           =   21
         Left            =   -71004
         TabIndex        =   52
         Top             =   2700
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label23"
         Height          =   228
         Index           =   22
         Left            =   -71004
         TabIndex        =   51
         Top             =   3024
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label24"
         Height          =   228
         Index           =   23
         Left            =   -71004
         TabIndex        =   50
         Top             =   3348
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label25"
         Height          =   228
         Index           =   24
         Left            =   -71004
         TabIndex        =   49
         Top             =   3672
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label26"
         Height          =   228
         Index           =   25
         Left            =   -71004
         TabIndex        =   48
         Top             =   3996
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label27"
         Height          =   228
         Index           =   26
         Left            =   -71004
         TabIndex        =   47
         Top             =   4320
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label28"
         Height          =   228
         Index           =   27
         Left            =   -71004
         TabIndex        =   46
         Top             =   4644
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label29"
         Height          =   228
         Index           =   28
         Left            =   -74568
         TabIndex        =   35
         Top             =   756
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label30"
         Height          =   228
         Index           =   29
         Left            =   -74568
         TabIndex        =   34
         Top             =   1080
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label31"
         Height          =   228
         Index           =   30
         Left            =   -74568
         TabIndex        =   33
         Top             =   1404
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label32"
         Height          =   228
         Index           =   31
         Left            =   -74568
         TabIndex        =   32
         Top             =   1728
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label33"
         Height          =   228
         Index           =   32
         Left            =   -74568
         TabIndex        =   31
         Top             =   2052
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label34"
         Height          =   228
         Index           =   33
         Left            =   -74568
         TabIndex        =   30
         Top             =   2376
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label35"
         Height          =   228
         Index           =   34
         Left            =   -74568
         TabIndex        =   29
         Top             =   2700
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label36"
         Height          =   228
         Index           =   35
         Left            =   -74568
         TabIndex        =   28
         Top             =   3024
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label37"
         Height          =   228
         Index           =   36
         Left            =   -74568
         TabIndex        =   27
         Top             =   3348
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label38"
         Height          =   228
         Index           =   37
         Left            =   -74568
         TabIndex        =   26
         Top             =   3672
         Width           =   1524
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   13
         Left            =   2808
         TabIndex        =   14
         Top             =   2376
         Width           =   1092
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   12
         Left            =   216
         TabIndex        =   13
         Top             =   2376
         Width           =   1308
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   11
         Left            =   5076
         TabIndex        =   12
         Top             =   1944
         Width           =   984
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   10
         Left            =   2808
         TabIndex        =   11
         Top             =   1944
         Width           =   1200
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   9
         Left            =   216
         TabIndex        =   10
         Top             =   1944
         Width           =   984
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   8
         Left            =   4104
         TabIndex        =   9
         Top             =   1512
         Width           =   1092
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   7
         Left            =   2376
         TabIndex        =   8
         Top             =   1512
         Width           =   876
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   6
         Left            =   216
         TabIndex        =   7
         Top             =   1512
         Width           =   1416
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   5
         Left            =   6480
         TabIndex        =   6
         Top             =   972
         Width           =   1092
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   4
         Left            =   216
         TabIndex        =   5
         Top             =   1188
         Width           =   1416
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   3
         Left            =   216
         TabIndex        =   4
         Top             =   864
         Width           =   1416
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   2
         Left            =   6480
         TabIndex        =   3
         Top             =   324
         Width           =   984
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   0
         Left            =   216
         TabIndex        =   2
         Top             =   540
         Width           =   1416
      End
      Begin VB.Label Lbl 
         Caption         =   "Label"
         Height          =   228
         Index           =   1
         Left            =   7668
         TabIndex        =   1
         Top             =   324
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   336
      Left            =   108
      Top             =   5400
      Visible         =   0   'False
      Width           =   1668
      _ExtentX        =   2942
      _ExtentY        =   593
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
      Caption         =   "Auxiliar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   336
      Left            =   1728
      Top             =   5400
      Visible         =   0   'False
      Width           =   1668
      _ExtentX        =   2942
      _ExtentY        =   593
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
      Caption         =   "Auxiliar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FModifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
  Carga_Comprobantes
  Leer_Clientes
End Sub

Private Sub Form_Load()
  ConectarAdodc AdoAux
  ConectarAdodc AdoClientes
End Sub

Public Sub Leer_Clientes()
  'Carga en el Data Combo los Clientes con su RUC
  sSQL = "SELECT * " _
       & "FROM Clientes " _
       & "WHERE Cliente <> '.' " _
       & "AND TD <>  'E' " _
       & "ORDER BY Cliente "
  SelectDBCombo DC(0), AdoClientes, sSQL, "Cliente"
End Sub

Sub Carga_Comprobantes()
  'Cargo el Comprobante
  sSQL = "SELECT Tipo_Comprobante_Codigo,Descripcion " _
       & "FROM Tipo_Comprobante " _
       & "WHERE Tipo_Comprobante_Codigo <> 0 "
  SelectDBCombo DC(2), AdoAux, sSQL, "Descripcion"
''  cod = .Fields("TipoComprobante")
''  If AdoTipoComprobante.Recordset.RecordCount > 0 Then
''     AdoTipoComprobante.Recordset.MoveFirst
''     AdoTipoComprobante.Recordset.Find ("Tipo_Comprobante_Codigo = '" & cod & "' ")
''     If Not AdoTipoComprobante.Recordset.EOF Then
''        DCTipoComprobante = AdoTipoComprobante.Recordset.Fields("Descripcion")
''     Else
''        MsgBox "El Comprobante no existe", vbInformation, "Aviso"
''     End If
''  End If
End Sub


