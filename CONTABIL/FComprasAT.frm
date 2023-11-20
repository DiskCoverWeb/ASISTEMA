VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FComprasAT 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPRAS"
   ClientHeight    =   7185
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   10560
   ForeColor       =   &H8000000F&
   Icon            =   "FComprasAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10560
   Begin VB.CommandButton CmdCerrar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cancelar"
      Height          =   750
      Left            =   9450
      Picture         =   "FComprasAT.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   103
      ToolTipText     =   "Salir"
      Top             =   840
      Width           =   960
   End
   Begin VB.Frame FrmRetencion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "RETENCIONES DE IVA POR:"
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
      TabIndex        =   2
      Top             =   105
      Width           =   9255
      Begin MSDataListLib.DataCombo DCRetISer 
         Bindings        =   "FComprasAT.frx":0AD8
         DataSource      =   "AdoRetIvaSerCC"
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Top             =   525
         Visible         =   0   'False
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CheckBox ChRetB 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Bienes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         Width           =   960
      End
      Begin VB.CheckBox ChRetS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Servicios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   5
         Top             =   525
         Width           =   1170
      End
      Begin MSDataListLib.DataCombo DCRetIBienes 
         Bindings        =   "FComprasAT.frx":0AF5
         DataSource      =   "AdoRetIvaBienesCC"
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
   End
   Begin VB.CommandButton CmdGrabar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Aceptar"
      Height          =   750
      Left            =   9450
      Picture         =   "FComprasAT.frx":0B15
      Style           =   1  'Graphical
      TabIndex        =   102
      ToolTipText     =   "Grabar"
      Top             =   105
      Width           =   960
   End
   Begin TabDlg.SSTab SSTCompras 
      Height          =   5400
      Left            =   105
      TabIndex        =   0
      Top             =   1680
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   9525
      _Version        =   393216
      TabHeight       =   420
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Comprobante de Compra"
      TabPicture(0)   =   "FComprasAT.frx":0E1F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label23"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DCSustento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OpcSi"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OpcNo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FraDctoModificado"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdAir"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Conceptos AIR"
      TabPicture(1)   =   "FComprasAT.frx":0E3B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Partidos Políticos"
      TabPicture(2)   =   "FComprasAT.frx":0E57
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton CmdAir 
         Caption         =   "&AIR"
         Height          =   444
         Left            =   9660
         Picture         =   "FComprasAT.frx":0E73
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Se ubica en la pestaña de Retenciones"
         Top             =   324
         Width           =   552
      End
      Begin VB.Frame Frame8 
         Caption         =   "SOLO PARTIDOS POLITICOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   -74895
         TabIndex        =   95
         Top             =   420
         Width           =   10050
         Begin VB.TextBox TxtMonTitGrat 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   7035
            TabIndex        =   101
            Text            =   "0.00"
            ToolTipText     =   "Se debe ingresar el valor de la transacción que corresponde al titulo oneroso, es decir, no oneroso para el informante"
            Top             =   2730
            Width           =   1905
         End
         Begin VB.TextBox TxtMonTitOner 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   7035
            TabIndex        =   99
            Text            =   "0.00"
            ToolTipText     =   "Se debe ingresar el valor de la transacción que corresponde al titulo oneroso, es decir, no gratuito para el informante"
            Top             =   1995
            Width           =   1905
         End
         Begin VB.TextBox TxtNumConParPol 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   7035
            MaxLength       =   10
            TabIndex        =   97
            Text            =   "0000000000"
            ToolTipText     =   $"FComprasAT.frx":1399
            Top             =   1260
            Width           =   1905
         End
         Begin MSForms.Label Label34 
            Height          =   330
            Left            =   945
            TabIndex        =   96
            Top             =   1260
            Width           =   4740
            Caption         =   "No. de contrato que sustenta la contratación"
            Size            =   "8361;582"
            FontName        =   "Arial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label35 
            Height          =   330
            Left            =   945
            TabIndex        =   98
            Top             =   1995
            Width           =   6000
            Caption         =   "Monto de Transacción que corresponde a Titulo Oneroso"
            Size            =   "10583;582"
            FontName        =   "Arial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label36 
            Height          =   330
            Left            =   945
            TabIndex        =   100
            Top             =   2730
            Width           =   6000
            Caption         =   "Monto de Transacción que corresponde a Titulo Gratuito"
            Size            =   "10583;582"
            FontName        =   "Arial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "PORCENTAJE DE LAS BASES IMPONIBLES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   105
         TabIndex        =   39
         Top             =   2940
         Width           =   4950
         Begin VB.TextBox TxtMontoIva 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   3465
            TabIndex        =   43
            Text            =   "0.00"
            ToolTipText     =   "Este valor se calcula automaticamente, es el resultado de aplicarle un porcentaje IVA a la Base Imponible gravada"
            Top             =   420
            Width           =   1275
         End
         Begin VB.TextBox TxtMontoIce 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3465
            TabIndex        =   47
            Top             =   945
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo DCPorcenIva 
            Bindings        =   "FComprasAT.frx":1454
            DataSource      =   "AdoPorIva"
            Height          =   315
            Left            =   945
            TabIndex        =   41
            ToolTipText     =   $"FComprasAT.frx":146C
            Top             =   420
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenIce 
            Bindings        =   "FComprasAT.frx":14FE
            DataSource      =   "AdoPorIce"
            Height          =   315
            Left            =   945
            TabIndex        =   45
            ToolTipText     =   $"FComprasAT.frx":1516
            Top             =   945
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSForms.Label Label9 
            Height          =   228
            Left            =   108
            TabIndex        =   40
            Top             =   420
            Width           =   636
            Caption         =   "% I.V.A."
            Size            =   "1122;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label14 
            Height          =   228
            Left            =   2160
            TabIndex        =   42
            Top             =   420
            Width           =   1260
            Caption         =   "Monto de I.V.A."
            Size            =   "2222;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label8 
            Height          =   228
            Left            =   108
            TabIndex        =   44
            Top             =   948
            Width           =   540
            Caption         =   "% ICE"
            Size            =   "952;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label12 
            Height          =   228
            Left            =   2160
            TabIndex        =   46
            Top             =   948
            Width           =   1152
            Caption         =   "Monto de ICE"
            Size            =   "2032;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "RETENCION DEL IVA POR BIENES Y/O SERVICIOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   5145
         TabIndex        =   48
         Top             =   2940
         Width           =   5055
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaServ 
            Bindings        =   "FComprasAT.frx":15A7
            DataSource      =   "AdoRetIvaServicios"
            Height          =   315
            Left            =   3150
            TabIndex        =   58
            ToolTipText     =   "Corresponde al porcentaje retenido en el IVA generado en la prestación de servicios"
            Top             =   735
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCPorcenRetenIvaBien 
            Bindings        =   "FComprasAT.frx":15C8
            DataSource      =   "AdoRetIvaBienes"
            Height          =   315
            Left            =   1470
            TabIndex        =   53
            ToolTipText     =   $"FComprasAT.frx":15E6
            Top             =   735
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
         End
         Begin VB.TextBox TxtIvaBienMonIva 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   1470
            MultiLine       =   -1  'True
            TabIndex        =   51
            Text            =   "FComprasAT.frx":1672
            ToolTipText     =   $"FComprasAT.frx":1677
            Top             =   420
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaBienValRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   1470
            TabIndex        =   55
            Top             =   1050
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaSerMonIva 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3150
            MultiLine       =   -1  'True
            TabIndex        =   57
            Text            =   "FComprasAT.frx":1716
            ToolTipText     =   $"FComprasAT.frx":171B
            Top             =   420
            Width           =   1590
         End
         Begin VB.TextBox TxtIvaSerValRet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   3150
            TabIndex        =   59
            Text            =   " "
            Top             =   1080
            Width           =   1590
         End
         Begin MSForms.Label Label22 
            Height          =   228
            Left            =   108
            TabIndex        =   54
            Top             =   1056
            Width           =   1164
            Caption         =   "Valor retenido"
            Size            =   "2053;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label20 
            Height          =   228
            Left            =   108
            TabIndex        =   50
            Top             =   420
            Width           =   1056
            Caption         =   "Monto de IVA"
            Size            =   "1863;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label19 
            Height          =   225
            Left            =   1470
            TabIndex        =   49
            Top             =   210
            Width           =   1485
            Caption         =   "IVA-BIENES"
            Size            =   "2619;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label18 
            Height          =   228
            Left            =   108
            TabIndex        =   52
            Top             =   732
            Width           =   1056
            Caption         =   "% Ret. IVA"
            Size            =   "1863;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label17 
            Height          =   225
            Left            =   3150
            TabIndex        =   56
            Top             =   210
            Width           =   1590
            Caption         =   "IVA-SERVICIOS"
            Size            =   "2805;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame FraDctoModificado 
         Caption         =   "NOTAS DE DEBITO/NOTAS DE CREDITO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   105
         TabIndex        =   60
         Top             =   4410
         Visible         =   0   'False
         Width           =   10095
         Begin VB.ComboBox CNumSerieTresComp 
            DataSource      =   "AdoAux"
            Height          =   288
            Left            =   6300
            TabIndex        =   66
            Top             =   420
            Width           =   1170
         End
         Begin VB.TextBox TxtNumSerieUnoComp 
            Height          =   330
            Left            =   5040
            MaxLength       =   3
            TabIndex        =   64
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   420
            Width           =   540
         End
         Begin VB.TextBox TxtNumSerieDosComp 
            Height          =   336
            Left            =   5676
            MaxLength       =   3
            TabIndex        =   65
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   420
            Width           =   540
         End
         Begin VB.TextBox TxtNumAutComp 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   8820
            MaxLength       =   10
            TabIndex        =   70
            ToolTipText     =   $"FComprasAT.frx":17B1
            Top             =   432
            Width           =   1170
         End
         Begin MSDataListLib.DataCombo DCDctoModif 
            Bindings        =   "FComprasAT.frx":183D
            DataSource      =   "AdoTipoComprobante"
            Height          =   288
            Left            =   108
            TabIndex        =   62
            ToolTipText     =   "Corresponde al tipo de comprobante que ha sido originalmente modificado antre la emisión de una nota de débito o crédito"
            Top             =   420
            Width           =   4848
            _ExtentX        =   8546
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
         End
         Begin MSMask.MaskEdBox MBFechaEmiComp 
            Height          =   330
            Left            =   7560
            TabIndex        =   68
            ToolTipText     =   $"FComprasAT.frx":185E
            Top             =   420
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "0"
         End
         Begin MSForms.Label Label30 
            Height          =   228
            Left            =   108
            TabIndex        =   61
            Top             =   216
            Width           =   1788
            Caption         =   "Documento Modificado"
            Size            =   "3154;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label31 
            Height          =   228
            Left            =   5040
            TabIndex        =   63
            Top             =   216
            Width           =   2316
            Caption         =   "No. Comprobante Modificado"
            Size            =   "4085;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label32 
            Height          =   225
            Left            =   7560
            TabIndex        =   67
            Top             =   210
            Width           =   1170
            Caption         =   "Fecha Emis."
            Size            =   "2064;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label33 
            Height          =   225
            Left            =   8820
            TabIndex        =   69
            Top             =   210
            Width           =   1065
            Caption         =   "No. Autoriz."
            Size            =   "1879;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "INGRESE LOS DATOS DE LA FACTURA, NOTA DE VENTA, ETC. ______________________ FORMULARIO 104"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1800
         Left            =   105
         TabIndex        =   16
         Top             =   1155
         Width           =   10095
         Begin VB.TextBox TxtNumAutor 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   8505
            MaxLength       =   10
            TabIndex        =   24
            Text            =   "0000000001"
            Top             =   540
            Width           =   1416
         End
         Begin MSMask.MaskEdBox MBFechaCad 
            Height          =   330
            Left            =   2625
            TabIndex        =   31
            ToolTipText     =   $"FComprasAT.frx":190A
            Top             =   1365
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin MSMask.MaskEdBox MBFechaRegis 
            Height          =   330
            Left            =   1365
            TabIndex        =   29
            ToolTipText     =   $"FComprasAT.frx":19C1
            Top             =   1365
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin VB.TextBox TxtBaseImpo 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   4515
            MaxLength       =   14
            MultiLine       =   -1  'True
            TabIndex        =   34
            Text            =   "FComprasAT.frx":1A49
            ToolTipText     =   "En este campo se debe ingresar el valor del comprobante cuya base imponible esta gravado con la tarifa del 0% de IVA"
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox TxtNumSerietres 
            Height          =   336
            Left            =   7455
            MaxLength       =   9
            TabIndex        =   22
            Text            =   "0000001"
            ToolTipText     =   $"FComprasAT.frx":1A50
            Top             =   525
            Width           =   960
         End
         Begin VB.TextBox TxtNumSerieDos 
            Height          =   336
            Left            =   6720
            MaxLength       =   3
            TabIndex        =   21
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   525
            Width           =   645
         End
         Begin VB.TextBox TxtNumSerieUno 
            Height          =   336
            Left            =   6090
            MaxLength       =   3
            TabIndex        =   20
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   525
            Width           =   645
         End
         Begin VB.TextBox TxtBaseImpoGrav 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6405
            MultiLine       =   -1  'True
            TabIndex        =   36
            Text            =   "FComprasAT.frx":1AF3
            ToolTipText     =   $"FComprasAT.frx":1AFA
            Top             =   1365
            Width           =   1380
         End
         Begin VB.TextBox TxtBaseImpoIce 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   8316
            MultiLine       =   -1  'True
            TabIndex        =   38
            Text            =   "FComprasAT.frx":1BA2
            ToolTipText     =   $"FComprasAT.frx":1BA7
            Top             =   1365
            Width           =   1380
         End
         Begin MSDataListLib.DataCombo DCTipoComprobante 
            Bindings        =   "FComprasAT.frx":1C39
            DataSource      =   "AdoTipoComp"
            Height          =   315
            Left            =   105
            TabIndex        =   18
            ToolTipText     =   $"FComprasAT.frx":1C5A
            Top             =   525
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSMask.MaskEdBox MBFechaEmi 
            Height          =   330
            Left            =   105
            TabIndex        =   27
            ToolTipText     =   $"FComprasAT.frx":1D02
            Top             =   1365
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin MSForms.Label Label15 
            Height          =   225
            Left            =   105
            TabIndex        =   25
            Top             =   945
            Width           =   3705
            Caption         =   " Fechas de:"
            Size            =   "6535;397"
            BorderStyle     =   1
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.Label Label7 
            Height          =   225
            Left            =   4515
            TabIndex        =   32
            Top             =   945
            Width           =   5160
            Caption         =   " Bases Imponibles de:"
            Size            =   "9102;397"
            BorderStyle     =   1
            FontName        =   "Arial"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.Label Label44 
            Height          =   228
            Left            =   4512
            TabIndex        =   33
            Top             =   1152
            Width           =   1272
            Caption         =   "Tarifa Cero %"
            Size            =   "2244;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label43 
            Height          =   228
            Left            =   108
            TabIndex        =   26
            Top             =   1152
            Width           =   744
            Caption         =   "Emisión"
            Size            =   "1312;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label3 
            Height          =   228
            Left            =   108
            TabIndex        =   17
            Top             =   312
            Width           =   1476
            Caption         =   "Tipo Comprobante"
            Size            =   "2603;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label4 
            Height          =   228
            Left            =   6096
            TabIndex        =   19
            Top             =   312
            Width           =   2112
            Caption         =   "No. de Serie y Secuencial"
            Size            =   "3725;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label6 
            Height          =   225
            Left            =   8505
            TabIndex        =   23
            Top             =   315
            Width           =   1380
            Caption         =   "No. Autorización"
            Size            =   "2434;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label10 
            Height          =   228
            Left            =   1476
            TabIndex        =   28
            Top             =   1152
            Width           =   744
            Caption         =   "Registro"
            Size            =   "1312;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label16 
            Height          =   228
            Left            =   2628
            TabIndex        =   30
            Top             =   1152
            Width           =   948
            Caption         =   "Caducidad"
            Size            =   "1672;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label11 
            Height          =   228
            Left            =   6408
            TabIndex        =   35
            Top             =   1152
            Width           =   1272
            Caption         =   "Gravada  I.V.A."
            Size            =   "2244;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label5 
            Height          =   228
            Left            =   8316
            TabIndex        =   37
            Top             =   1152
            Width           =   420
            Caption         =   "ICE"
            Size            =   "741;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.OptionButton OpcNo 
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   4305
         TabIndex        =   13
         ToolTipText     =   $"FComprasAT.frx":1DAE
         Top             =   315
         Value           =   -1  'True
         Width           =   636
      End
      Begin VB.OptionButton OpcSi 
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   2940
         TabIndex        =   12
         ToolTipText     =   $"FComprasAT.frx":1E46
         Top             =   315
         Width           =   636
      End
      Begin VB.Frame Frame2 
         Caption         =   "INGRESE LOS DATOS DE LA RETENCION _________________________________________ FORMULARIO 103"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   4965
         Left            =   -74892
         TabIndex        =   72
         Top             =   315
         Width           =   10155
         Begin MSDataListLib.DataCombo DCRetFuente 
            Bindings        =   "FComprasAT.frx":1EDE
            DataSource      =   "AdoRetFuente"
            Height          =   315
            Left            =   2415
            TabIndex        =   74
            Top             =   315
            Visible         =   0   'False
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.CheckBox ChRetF 
            Caption         =   "Retención en la Fuente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   73
            Top             =   315
            Visible         =   0   'False
            Width           =   2328
         End
         Begin VB.TextBox TxtValConA 
            Enabled         =   0   'False
            Height          =   336
            Left            =   8715
            TabIndex        =   90
            Top             =   1470
            Width           =   1275
         End
         Begin VB.TextBox TxtPorRetConA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   8085
            TabIndex        =   88
            Top             =   1470
            Width           =   645
         End
         Begin VB.TextBox TxtTotalReten 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   8715
            TabIndex        =   93
            Text            =   "0.00"
            ToolTipText     =   "Sumatoria total de las retenciones"
            Top             =   4515
            Width           =   1275
         End
         Begin VB.TextBox TxtSumatoria 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Left            =   8085
            MultiLine       =   -1  'True
            TabIndex        =   82
            Text            =   "FComprasAT.frx":1EF9
            Top             =   735
            Width           =   1905
         End
         Begin VB.TextBox TxtBimpConA 
            Alignment       =   1  'Right Justify
            Height          =   336
            Left            =   6720
            TabIndex        =   86
            Top             =   1470
            Width           =   1380
         End
         Begin VB.TextBox TxtNumUnoAutComRet 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   2415
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   80
            ToolTipText     =   $"FComprasAT.frx":1F00
            Top             =   840
            Width           =   1590
         End
         Begin VB.TextBox TxtNumTresComRet 
            Height          =   336
            Left            =   1365
            MaxLength       =   9
            TabIndex        =   78
            Text            =   "0000001"
            ToolTipText     =   $"FComprasAT.frx":1F8C
            Top             =   840
            Width           =   960
         End
         Begin VB.TextBox TxtNumDosComRet 
            Height          =   336
            Left            =   735
            MaxLength       =   3
            TabIndex        =   77
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al punto dde emisión"
            Top             =   840
            Width           =   540
         End
         Begin VB.TextBox TxtNumUnoComRet 
            Height          =   336
            Left            =   105
            MaxLength       =   3
            TabIndex        =   76
            Text            =   "001"
            ToolTipText     =   "En este campo se debe ingresar el número de serie del comprobante, la parte correspondiente al código del establecimiento"
            Top             =   840
            Width           =   540
         End
         Begin MSDataListLib.DataCombo DCConceptoRet 
            Bindings        =   "FComprasAT.frx":202E
            DataSource      =   "AdoConceptoRet"
            Height          =   315
            Left            =   105
            TabIndex        =   84
            Top             =   1470
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSDataGridLib.DataGrid DGConceptoAir 
            Bindings        =   "FComprasAT.frx":204B
            Height          =   2595
            Left            =   105
            TabIndex        =   91
            Top             =   1890
            Width           =   9945
            _ExtentX        =   17542
            _ExtentY        =   4577
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   19
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Datos Ingresados"
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
                  LCID            =   3082
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
                  LCID            =   3082
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
         Begin MSForms.Label Label42 
            Height          =   225
            Left            =   7035
            TabIndex        =   92
            Top             =   4515
            Width           =   1695
            Caption         =   "Total Retenciones:"
            Size            =   "2990;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label37 
            Height          =   225
            Left            =   5460
            TabIndex        =   81
            Top             =   810
            Width           =   2535
            Caption         =   "Base Imponible de Retención:"
            Size            =   "4471;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   2
         End
         Begin MSForms.Label Label21 
            Height          =   225
            Left            =   105
            TabIndex        =   75
            Top             =   630
            Width           =   2220
            Caption         =   "No. de Serie y Secuencial"
            Size            =   "3916;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label28 
            Height          =   225
            Left            =   2415
            TabIndex        =   79
            Top             =   630
            Width           =   1590
            Caption         =   "No. de Autorización"
            Size            =   "2794;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label24 
            Height          =   225
            Left            =   8715
            TabIndex        =   89
            Top             =   1260
            Width           =   1275
            Caption         =   "Valor Retenido"
            Size            =   "2249;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label25 
            Height          =   225
            Left            =   6720
            TabIndex        =   85
            Top             =   1260
            Width           =   1275
            Caption         =   "Base Imponible"
            Size            =   "2244;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label26 
            Height          =   225
            Left            =   8085
            TabIndex        =   87
            Top             =   1260
            Width           =   645
            Caption         =   "% Ret."
            Size            =   "1138;397"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label27 
            Height          =   228
            Left            =   108
            TabIndex        =   83
            Top             =   1260
            Width           =   4908
            Caption         =   "RETENCION EN LA FUENTE DEL IMPUESTO A  LA RENTA "
            Size            =   "8657;402"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSDataListLib.DataCombo DCSustento 
         Bindings        =   "FComprasAT.frx":2067
         DataSource      =   "AdoSustento"
         Height          =   315
         Left            =   105
         TabIndex        =   15
         ToolTipText     =   "En este campo de selección se despliega un lista de tipos de sustentos tributarios correspondientes a la transacción escogida"
         Top             =   840
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSForms.Label Label2 
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   315
         Width           =   1485
         Caption         =   "Devolución I.V.A.:"
         Size            =   "2619;397"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label1 
         Height          =   228
         Left            =   108
         TabIndex        =   14
         Top             =   636
         Width           =   6204
         Caption         =   "Identifique el tipo de sustento tributario que le  corresponde a esta transacción:"
         Size            =   "10943;402"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label23 
         Height          =   225
         Left            =   420
         TabIndex        =   94
         Top             =   2310
         Width           =   1905
         Caption         =   "Devolución de I.V.A."
         Size            =   "3360;397"
         FontName        =   "Arial"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSAdodcLib.Adodc AdoSustento 
      Height          =   330
      Left            =   2730
      Top             =   3045
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoSustento"
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
   Begin MSAdodcLib.Adodc AdoTipoIdentificacion 
      Height          =   330
      Left            =   210
      Top             =   2100
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoTipoIden"
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
   Begin MSAdodcLib.Adodc AdoTipoComprobante 
      Height          =   330
      Left            =   210
      Top             =   2415
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoTipoComp"
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
   Begin MSAdodcLib.Adodc AdoRetIvaBienes 
      Height          =   330
      Left            =   210
      Top             =   2730
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoRetenBienes"
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
   Begin MSAdodcLib.Adodc AdoRetIvaServicios 
      Height          =   330
      Left            =   210
      Top             =   3045
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoRetenServicios"
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
   Begin MSAdodcLib.Adodc AdoPorIva 
      Height          =   330
      Left            =   2730
      Top             =   2730
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoPorIva"
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
   Begin MSAdodcLib.Adodc AdoPorIce 
      Height          =   330
      Left            =   2730
      Top             =   2415
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoPorIce"
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
   Begin MSAdodcLib.Adodc AdoCaTrTiCom 
      Height          =   330
      Left            =   210
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Catal.Tributarios y Tipos de Comprobantes"
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
   Begin MSAdodcLib.Adodc AdoConceptoRet 
      Height          =   330
      Left            =   210
      Top             =   3675
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoConceptoRet"
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
   Begin MSAdodcLib.Adodc AdoTransCompras 
      Height          =   330
      Left            =   210
      Top             =   3990
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "TransCompras"
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
   Begin MSAdodcLib.Adodc AdoAsientoAir 
      Height          =   330
      Left            =   210
      Top             =   4305
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AsientoAir"
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
   Begin MSAdodcLib.Adodc AdoClientes 
      Height          =   330
      Left            =   2730
      Top             =   2100
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Clientes"
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
   Begin MSAdodcLib.Adodc AdoRetIvaSerCC 
      Height          =   330
      Left            =   2730
      Top             =   3675
      Visible         =   0   'False
      Width           =   3270
      _ExtentX        =   5768
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
      Caption         =   "RetencionFuenteServicios"
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
   Begin MSAdodcLib.Adodc AdoRetIvaBienesCC 
      Height          =   330
      Left            =   2730
      Top             =   3990
      Visible         =   0   'False
      Width           =   3270
      _ExtentX        =   5768
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
      Caption         =   "RetencionFuenteBienes"
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
   Begin MSAdodcLib.Adodc AdoAsientoCompras 
      Height          =   330
      Left            =   210
      Top             =   4935
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AsientoCompras"
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
   Begin MSAdodcLib.Adodc AdoRetFuente 
      Height          =   330
      Left            =   210
      Top             =   5250
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "RetencionFuente"
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
   Begin MSAdodcLib.Adodc AdoTransAir 
      Height          =   330
      Left            =   210
      Top             =   4620
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "TransAir"
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
      Left            =   2730
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc AdoAsientos 
      Height          =   330
      Left            =   210
      Top             =   5565
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Asientos"
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
   Begin VB.Label LblNumIdent 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   330
      Left            =   7560
      TabIndex        =   10
      Top             =   1260
      Width           =   1800
   End
   Begin VB.Label LblTD 
      Alignment       =   2  'Center
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
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   7245
      TabIndex        =   8
      Top             =   1260
      Width           =   330
   End
   Begin VB.Label LblProveedor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente/Proveedor"
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
      TabIndex        =   104
      Top             =   1260
      Width           =   7155
   End
   Begin MSForms.Label Label41 
      Height          =   225
      Left            =   105
      TabIndex        =   7
      Top             =   1050
      Width           =   7155
      BackColor       =   16761024
      Caption         =   "Proveedor"
      Size            =   "12621;397"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label13 
      Height          =   225
      Left            =   7560
      TabIndex        =   9
      Top             =   1050
      Width           =   1695
      BackColor       =   16761024
      Caption         =   "No. de Identificación"
      Size            =   "2984;402"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LblMensaje 
      Height          =   900
      Left            =   210
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   9120
      Caption         =   "ANEXOS TRANSACCIONALES"
      Size            =   "16087;1587"
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   585
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "FComprasAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MBFecha As MaskEdBox
Dim FechaRegis As Date
Dim OP As Boolean
Dim cod, x, Rb, Rf, rs, CapDm As Byte
Dim SumAnio, Aniocad, AniocadAux, CodPorIva, CodRetBien, CodRetServ As Integer
Dim CalmIva, CalmIce, CalIbMi, CalIsMi, ac, SUM, cal As Double
Dim CapDcto, Cap, Cap1, Ct, ValorP, AuxCodUs, Opc, conta, ch, Ch1, CodSus, Bien, Serv, CargaTC, CodPorIce As String
Dim Espizq, Espder, Captc, PorIce, PorIva, CodProv, CodProv1, NumCed, Mifecha  As String

Private Sub ChRetB_Click()
    If ChRetB.value <> 0 Then
       ch = 1
       Ch1 = "B"
       DCRetIBienes.Visible = True
       TxtIvaBienMonIva.Enabled = True
       DCPorcenRetenIvaBien.Enabled = True
       TxtIvaBienValRet.Enabled = True
    Else
       DCRetIBienes.Visible = False
       TxtIvaBienMonIva.Enabled = False
       DCPorcenRetenIvaBien.Enabled = False
       TxtIvaBienValRet.Enabled = False
       ch = 1
       Ch1 = "S"
    End If
    If ChRetB.value <> 0 And ChRetS.value <> 0 Then
       Ch1 = "X"
    End If
End Sub

Private Sub ChRetF_Click()
  If ChRetF.value = 0 Then DCRetFuente.Visible = False Else DCRetFuente.Visible = True
End Sub

Private Sub ChRetS_Click()
    If ChRetS.value <> 0 Then
       ch = 1
       Ch1 = "S"
       DCRetISer.Visible = True
       TxtIvaSerMonIva.Enabled = True
       DCPorcenRetenIvaServ.Enabled = True
       TxtIvaSerValRet.Enabled = True
    Else
       DCRetISer.Visible = False
       TxtIvaSerMonIva.Enabled = False
       DCPorcenRetenIvaServ.Enabled = False
       TxtIvaSerValRet.Enabled = False
    End If
    If ChRetB.value <> 0 And ChRetS.value <> 0 Then
       Ch1 = "X"
    End If
End Sub

Private Sub CmdAir_Click()
  SSTCompras.Tab = 1
  TxtNumUnoComRet.SetFocus
End Sub

Private Sub CmdCerrar_Click()
    Total_Ret = 0
    Total_RetIVA = 0
   'Borra Asiento Compras
    sSQL = "DELETE * " _
         & "FROM Asiento_Compras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    ConectarAdoExecute sSQL
   'Borra Asiento Air
    sSQL = "DELETE * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND Tipo_Trans = 'C' " _
         & "AND T_No = " & Trans_No & " "
    ConectarAdoExecute sSQL
    Unload FComprasAT
End Sub

Private Sub CmdGrabar_Click()
   'Valido por si acaso exista algun valor con 0
    TextoValido TxtIvaBienMonIva, True, , 2
    TextoValido TxtBaseImpo, True, , 2
    TextoValido TxtBaseImpoGrav, True, , 2
    TextoValido TxtBaseImpoIce, True, , 2
    TextoValido TxtMontoIva, True, , 2
    TextoValido TxtMontoIce, True, , 2
    TextoValido TxtIvaBienMonIva, True, , 2
    TextoValido TxtIvaBienValRet, True, , 2
    TextoValido TxtIvaSerMonIva, True, , 2
    TextoValido TxtIvaSerValRet, True, , 2
    Grabacion
    Total_Ret = 0
    Total_RetIVA = 0
    sSQL = "SELECT * " _
         & "FROM Asiento " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    SelectAdodc AdoAsientos, sSQL
    OpcTM = 1
    OpcDH = 2
    NoCheque = Ninguno
   'Grabamos el Asiento de la Compra
    sSQL = "SELECT * " _
         & "FROM Asiento_Compras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    SelectAdodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
        'Porcentaje por Servicio: 0,30,100
         Cta = .Fields("Cta_Servicio")
         DetalleComp = "Retencion del " & .Fields("Porc_Bienes") & "%, Factura No. " & .Fields("Secuencial") & ", de " & NombreCliente
         LeerCta Cta
         ValorDH = .Fields("ValorRetServicios")
         Total_RetIVA = Total_RetIVA + .Fields("ValorRetServicios")
         If ValorDH > 0 Then InsertarAsiento AdoAsientos
        'Porcentaje por Bienes: 0,70,100
         Cta = .Fields("Cta_Bienes")
         DetalleComp = "Retencion del " & .Fields("Porc_Servicios") & "%, Factura No. " & .Fields("Secuencial") & ", de " & NombreCliente
         LeerCta Cta
         ValorDH = .Fields("ValorRetBienes")
         Total_RetIVA = Total_RetIVA + .Fields("ValorRetBienes")
         If ValorDH > 0 Then InsertarAsiento AdoAsientos
     End If
    End With
   'Grabamos el Asiento de las Retenciones
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'C' " _
         & "ORDER BY Cta_Retencion,A_No,ValRet "
    SelectAdodc AdoAux, sSQL
    With AdoAux.Recordset
     If .RecordCount > 0 Then
         Do While Not .EOF
            Cta = .Fields("Cta_Retencion")
            DetalleComp = "Retencion (" & .Fields("CodRet") & ") No. " & .Fields("SecRetencion") & " del " & (.Fields("Porcentaje") * 100) & "%, de " & NombreCliente
            LeerCta Cta
            ValorDH = .Fields("ValRet")
            Total_Ret = Total_Ret + .Fields("ValRet")
            If ValorDH > 0 Then InsertarAsiento AdoAsientos
           .MoveNext
         Loop
     End If
    End With
    DetalleComp = Ninguno
    Unload FComprasAT
End Sub

Private Sub DCConceptoRet_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub DCConceptoRet_LostFocus()
    OP = False
    If IsNumeric(DCConceptoRet.Text) Then
       MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCConceptoRet.SetFocus
    Else
       With AdoConceptoRet.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Codigo = '" & SinEspaciosIzq(DCConceptoRet) & "' ")
            If Not .EOF Then
               TxtPorRetConA = .Fields("Porcentaje")
               If .Fields("Ingresar_Porcentaje") = "S" Then OP = True
              'MsgBox .Fields("Porcentaje")
            Else
               MsgBox "No encontro este código vuelva a buscar"
            End If
        End If
       End With
       TxtBimpConA = TxtSumatoria
    End If
    If OP Then
       TxtPorRetConA.Enabled = True
       TxtPorRetConA.SetFocus
    Else
       TxtPorRetConA.Enabled = False
    End If
End Sub

Private Sub DCDctoModif_LostFocus()
    Captura_TipoComprobante_DctoModificado
End Sub

Private Sub DCPorcenIva_GotFocus()
    MarcarTexto DCPorcenIva
End Sub

Private Sub DCPorcenIva_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub DCTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub DCTipoComprobante_LostFocus()
    If IsNumeric(DCTipoComprobante.Text) Then
       MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCTipoComprobante.Text = ""
       Carga_TipoComprobante (DCSustento)
       DCTipoComprobante.SetFocus
       Captura_TipoComprobante
    Else
       Captura_TipoComprobante
    End If
End Sub

Private Sub DGConceptoAir_KeyDown(KeyCode As Integer, Shift As Integer)
 Keys_Especiales Shift
 If CtrlDown And KeyCode = vbKeyDelete Then
   Titulo = "Aviso"
   Mensajes = "Desea Eliminar la Retención"
   If BoxMensaje = vbYes Then
      With AdoAsientoAir.Recordset
       If .RecordCount > 0 Then
           Codigo = .Fields("CodRet")
           No_Desde = .Fields("SecRetencion")
           Mifecha = BuscarFecha(.Fields("FechaEmiRet"))
           Codigo1 = .Fields("AutRetencion")
           J = .Fields("A_No")
           sSQL = "DELETE * " _
                & "FROM Asiento_Air " _
                & "WHERE Item = '" & NumEmpresa & "' " _
                & "AND CodigoU = '" & CodigoUsuario & "' " _
                & "AND IdProv = '" & CodigoCliente & "' " _
                & "AND T_No = " & Trans_No & " " _
                & "AND Tipo_Trans = 'C' " _
                & "AND A_No = " & J & " " _
                & "AND CodRet = '" & Codigo & "' "
           ConectarAdoExecute sSQL
       End If
       AdoAsientoAir.Refresh
       Calculo_Sumatoria
     End With
   End If
 End If
End Sub

Private Sub MBFechaCad_GotFocus()
    MarcarTexto MBFechaCad
End Sub

Private Sub MBFechaCad_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmi_GotFocus()
    MarcarTexto MBFechaEmi
End Sub

Private Sub MBFechaEmi_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmiComp_GotFocus()
    MarcarTexto MBFechaEmiComp
End Sub

Private Sub MBFechaEmiComp_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub MBFechaEmiComp_LostFocus()
    FechaValida MBFechaEmiComp
End Sub

Private Sub OpcNo_LostFocus()
    If OpcNo.value = True Then ValorP = "N"
End Sub

Private Sub OpcSi_LostFocus()
    If OpcSi.value = True Then ValorP = "S"
End Sub

Private Sub SSTCompras_Click(PreviousTab As Integer)
    Select Case PreviousTab
        Case 0: If ChRetF.Visible Then ChRetF.SetFocus
        Case 1: OpcSi.SetFocus
    End Select
End Sub

Private Sub TxtBaseImpo_GotFocus()
    MarcarTexto TxtBaseImpo
End Sub

Private Sub TxtBaseImpo_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpo_LostFocus()
    TextoValido TxtBaseImpo, True, , 0
End Sub

Private Sub TxtBaseImpoGrav_GotFocus()
    MarcarTexto TxtBaseImpoGrav
End Sub

Private Sub TxtBaseImpoGrav_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoGrav_LostFocus()
    TextoValido TxtBaseImpoGrav, True, , 0
End Sub

Private Sub TxtBaseImpoIce_GotFocus()
    MarcarTexto TxtBaseImpoIce
End Sub

Private Sub TxtBaseImpoIce_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtBaseImpoIce_LostFocus()
    TextoValido TxtBaseImpoIce, True, , 0
End Sub

Private Sub TxtBimpConA_GotFocus()
   MarcarTexto TxtBimpConA
End Sub

Private Sub TxtBimpConA_LostFocus()
    TextoValido TxtBimpConA, True, , 2
    TextoValido TxtSumatoria, True, , 2
    RatonNormal
   'Valida que la base imponible no sea mayor que la BIG y la BIcero
    If CTNumero(TxtBimpConA, 2) > CTNumero(TxtSumatoria, 2) Then
       MsgBox "La Base Imponible debe ser menor o igual a la " & vbCrLf _
       & "Base Imponible Gravada + la Base Imponible 0%", vbInformation, "Aviso"
       TxtBimpConA.Text = 0
       TxtBimpConA.SetFocus
    Else
       If Not OP Then
          TxtValConA = CTNumero(TxtBimpConA, 2) * (CTNumero(TxtPorRetConA, 2) / 100)
          Insertar_DataGrid
          If (cod = 4) Or (cod = 5) Then
             DCDctoModif.SetFocus
          Else
             TxtNumConParPol.SetFocus
          End If
       End If
    End If
End Sub

Sub Insertar_DataGrid()
    'Selecciona el numero mayor para continuar la secuencia en el
    'campo T_No y A_No
    If Val(CCur(TxtBimpConA)) > 0 Then
       RatonReloj
       Espizq = SinEspaciosIzq(DCConceptoRet)
       Espder = Trim(Mid(DCConceptoRet, Len(Espizq) + 3, Len(DCConceptoRet)))
       SetAdoAddNew "Asiento_Air"
       SetAdoFields "CodRet", Espizq
       SetAdoFields "Detalle", Espder
       SetAdoFields "BaseImp", CTNumero(TxtBimpConA, 2)
       SetAdoFields "Porcentaje", CTNumero(TxtPorRetConA, 2) / 100
       SetAdoFields "ValRet", CTNumero(TxtValConA, 2)
       SetAdoFields "EstabRetencion", TxtNumUnoComRet
       SetAdoFields "PtoEmiRetencion", TxtNumDosComRet
       SetAdoFields "SecRetencion", CTNumero(TxtNumTresComRet)
       SetAdoFields "AutRetencion", TxtNumUnoAutComRet
       SetAdoFields "FechaEmiRet", MBFechaRegis
       SetAdoFields "Cta_Retencion", SinEspaciosIzq(DCRetFuente)
       SetAdoFields "EstabFactura", TxtNumSerieUno
       SetAdoFields "PuntoEmiFactura", TxtNumSerieDos
       SetAdoFields "Factura_No", CTNumero(TxtNumSerietres)
       SetAdoFields "IdProv", CodigoCliente
       SetAdoFields "A_No", Maximo_De("Asiento_Air", "A_No")
       SetAdoFields "T_No", Trans_No
       SetAdoFields "Tipo_Trans", "C"
       SetAdoUpdate
              
      'Despliega los datos en el DataGrid
       sSQL = "SELECT * " _
            & "FROM Asiento_Air " _
            & "WHERE CodRet <> '.' " _
            & "AND Item = '" & NumEmpresa & "' " _
            & "AND CodigoU = '" & CodigoUsuario & "' " _
            & "AND T_No = " & Trans_No & " " _
            & "AND Tipo_Trans = 'C' " _
            & "ORDER BY CodRet "
       SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL, "Sustento"
         
      'Se situa en el combo de retención AIR
       If ChRetF.Visible Then DCRetFuente.SetFocus Else TxtNumUnoComRet.SetFocus
       
      'Realiza la Sumatoria de las Retenciones
       ac = ac + TxtValConA
       TxtTotalReten = ac
    End If
    RatonNormal
End Sub

Private Sub DCPorcenIce_LostFocus()
    If Not IsNumeric(DCPorcenIce) Then
       MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCPorcenIce = ""
       'Carga_PorcentajeIce
       DCPorcenIce.SetFocus
    Else
       'Busca y captura el codigo de Porcentaje IVA
       PorIce = (DCPorcenIce.Text)
       With AdoPorIce.Recordset
            If .RecordCount > 0 Then
               .MoveFirst
               .Find ("Porc = '" & PorIce & "' ")
               If Not .EOF Then
                  CodPorIce = .Fields("Codigo")
               Else
                  'MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
               End If
             End If
       End With
        
       Total_IVA = 0
       Total_IVA = CTNumero(TxtBaseImpoIce, 2)
       TxtMontoIce = 0
      'Calcula el Porcentaje de Ice
       CalIbMi = (Total_IVA * DCPorcenIce) / 100
       TxtMontoIce = CalIbMi
    End If
    
    'Coloca el valor de Monto IVA dependiendo si se activo Bienes o Servicios
    If ChRetB + ChRetS = 0 Then
       TxtIvaBienMonIva = TxtMontoIva
    End If
    If ChRetB.value <> 0 Then
       TxtIvaBienMonIva = TxtMontoIva
       TxtIvaSerMonIva = 0
    Else
       If ChRetS.value <> 0 Then
          TxtIvaSerMonIva = TxtMontoIva
          TxtIvaBienMonIva = 0
       End If
    End If
End Sub

Private Sub DCPorcenIva_LostFocus()
    If Not IsNumeric(DCPorcenIva) Then
       MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCPorcenIva = ""
       'Carga_PorcentajeIva (MBFechaRegis)
       DCPorcenIva.SetFocus
    Else
       'Busca y captura el codigo de Porcentaje IVA
       PorIva = SinEspaciosDer(DCPorcenIva.Text)
       CodPorIva = "0"
       With AdoPorIva.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Porc = '" & PorIva & "' ")
            If Not .EOF Then CodPorIva = .Fields("Codigo")
        End If
       End With
       Total_IVA = 0
       Total_IVA = CTNumero(TxtBaseImpoGrav, 2)
      'Calcula el Porcentaje de Iva
       CalmIva = (Total_IVA * DCPorcenIva) / 100
       TxtMontoIva = CalmIva
    End If
End Sub

Private Sub DCPorcenRetenIvaBien_LostFocus()
    CodRetBien = 0
    If Not IsNumeric(DCPorcenRetenIvaBien) Then
       MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCPorcenRetenIvaBien = ""
       Carga_RetencionIvaBienes_Servicios
       DCPorcenRetenIvaBien.SetFocus
    Else
       'Busca y captura el codigo de Porcentaje retencion Iva Bienes
       With AdoRetIvaBienes.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaBien) & " ")
            If Not .EOF Then CodRetBien = .Fields("Codigo")
        Else
            MsgBox "Código incorrecto", vbInformation, "Aviso"
        End If
       End With
       Total_IVA = CTNumero(TxtIvaBienMonIva, 2)
      'Calcula la retencion Iva Bienes
       CalIbMi = (Total_IVA * CInt(DCPorcenRetenIvaBien)) / 100
       TxtIvaBienValRet = CalIbMi
    End If
    TxtIvaSerMonIva = Format(CTNumero(TxtMontoIva, 2) - CTNumero(TxtIvaBienMonIva, 2), "#,##0.00")
End Sub

Private Sub DCPorcenRetenIvaServ_LostFocus()
    CodRetServ = 0
   'Activo el casillero para que ingrese el valor si el porcentaje es 70/100
    If DCPorcenRetenIvaServ = "70/100" Then
       Ct = "Si"
       TxtIvaSerValRet.Text = ""
       TxtIvaSerValRet.Enabled = True
       TxtIvaSerValRet.SetFocus
    Else
      If Not IsNumeric(DCPorcenRetenIvaServ) Then
         MsgBox "No ingrese caracteres. Vuelva a seleccionar.", vbInformation, "Aviso"
         DCPorcenRetenIvaServ = ""
         Carga_RetencionIvaBienes_Servicios
         DCPorcenRetenIvaServ.SetFocus
      End If
    End If
    
    'Busca captura el codigo de Porcentaje retencion Iva Servicios
    With AdoRetIvaServicios.Recordset
     If .RecordCount > 0 Then
        .MoveFirst
        .Find ("Porc = " & SinEspaciosDer(DCPorcenRetenIvaServ) & " ")
         If Not .EOF Then CodRetServ = .Fields("Codigo")
     Else
        MsgBox "Código Incorrecto", vbInformation, "Aviso"
     End If
    End With
    Ct = "No"
    Total_IVA = 0
    Total_IVA = CTNumero(TxtIvaSerMonIva, 2)
    If DCPorcenRetenIvaServ = "70/100" Then
    Else
       CalIsMi = (Total_IVA * CInt(DCPorcenRetenIvaServ)) / 100
       TxtIvaSerValRet = CalIsMi
       TxtIvaSerValRet.Enabled = False
    End If
    SSTCompras.Tab = 0
    SSTCompras.SetFocus
End Sub

Private Sub DCSustento_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub DCSustento_LostFocus()
    If IsNumeric(DCSustento.Text) Then
       MsgBox "No ingrese números. Vuelva a seleccionar.", vbInformation, "Aviso"
       DCSustento.Text = ""
       Carga_CreditoTributario
       DCSustento.SetFocus
       Carga_TipoComprobante (SinEspaciosIzq(DCSustento))
    Else
       Carga_TipoComprobante (SinEspaciosIzq(DCSustento))
    End If
End Sub

Private Sub Form_Activate()
   Carga_Datos_Iniciales MBFecha, Nuevo
   LblTD.Caption = TipoBenef                  ' Tipo de Cliente: C,R,P,O
   LblNumIdent = CICliente                    ' CI o RUC del Cliente
   Label41.Caption = "PROVEEDOR: " & TipoContribuyente
   LblProveedor.Caption = " " & NombreCliente ' Nombre del Cliente
   MBFechaEmi = FechaComp
   MBFechaRegis = FechaComp
   MBFechaCad = FechaComp
   TxtNumSerietres = "0000001"
   TxtNumSerieUno = "001"
   TxtNumSerieDos = "001"
   TxtNumAutor = String(10, "0")
   TxtNumUnoComRet = "001"
   TxtNumDosComRet = "001"
   TxtNumTresComRet = "0000001"
   TxtNumUnoAutComRet = String(10, "0")
  'CodigoCliente
  'Ultima Factura del Proveedor
   sSQL = "SELECT TOP 1 * " _
        & "FROM Trans_Compras " _
        & "WHERE IdProv = '" & CodigoCliente & "' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "ORDER BY Fecha DESC,Secuencial DESC "
   SelectAdodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      TxtNumSerietres = AdoAux.Recordset.Fields("Secuencial") + 1
      MBFechaCad = AdoAux.Recordset.Fields("FechaCaducidad")
      TxtNumSerieUno = AdoAux.Recordset.Fields("Establecimiento")
      TxtNumSerieDos = AdoAux.Recordset.Fields("PuntoEmision")
      TxtNumAutor = AdoAux.Recordset.Fields("Autorizacion")
   Else
      TxtNumAutor = Autorizacion
   End If
  'Ultima Retencion Emitida
   TxtNumUnoComRet = "001"
   TxtNumDosComRet = "001"
   TxtNumTresComRet = 1
   TxtNumUnoAutComRet = "1234567890"
   sSQL = "SELECT TOP 1 * " _
        & "FROM Trans_Air " _
        & "WHERE Tipo_Trans = 'C' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND Fecha <= #" & BuscarFecha(MBFechaEmi) & "# " _
        & "AND Porcentaje > 0 " _
        & "ORDER BY Fecha DESC,SecRetencion DESC "
   SelectAdodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      TxtNumUnoComRet = AdoAux.Recordset.Fields("EstabRetencion")
      TxtNumDosComRet = AdoAux.Recordset.Fields("PtoEmiRetencion")
      TxtNumTresComRet = AdoAux.Recordset.Fields("SecRetencion") + 1
      TxtNumUnoAutComRet = AdoAux.Recordset.Fields("AutRetencion")
   Else
      TxtNumUnoAutComRet = AutorizaRet
   End If
   'MsgBox sSQL & vbCrLf & AdoAux.Recordset.Fields("SecRetencion")
End Sub

Private Sub Form_Load()
   CentrarForm FComprasAT
   ConectarAdodc AdoAux
   ConectarAdodc AdoSustento
   ConectarAdodc AdoTipoIdentificacion
   ConectarAdodc AdoTipoComprobante
   ConectarAdodc AdoRetIvaBienes
   ConectarAdodc AdoRetIvaServicios
   ConectarAdodc AdoPorIce
   ConectarAdodc AdoPorIva
   ConectarAdodc AdoConceptoRet
   ConectarAdodc AdoAsientoCompras
   ConectarAdodc AdoTransCompras
   ConectarAdodc AdoAsientoAir
   ConectarAdodc AdoTransAir
   ConectarAdodc AdoAsientos
   ConectarAdodc AdoClientes
   ConectarAdodc AdoRetFuente
   ConectarAdodc AdoRetIvaSerCC
   ConectarAdodc AdoRetIvaBienesCC
End Sub

Private Sub MBFechaCad_LostFocus()
   'Verifico que la fecha de caducidad no sea mayor a la de emisión
   FechaValida MBFechaCad
   If MBFechaCad = "00/00/0000" Then
      MsgBox "Fecha no válida, vuelva a ingresar", vbInformation, "Aviso"
      MBFechaCad.SetFocus
   Else
        'Captura el año de la fecha de emisión
        Anio = Year(MBFechaEmi)
        SumAnio = Anio + 1  'Emisión + 1 año
        Aniocad = Year(MBFechaCad)
        AniocadAux = Aniocad + 1 'Asigno en otra variable el año de caducidad
        'Verifica si el año de caducidad es menor que el año de Emisión
        If (Aniocad < Anio) Then
           MsgBox "La Fecha de Caducidad no debe ser < a la Fecha de Emisión", vbInformation, "Aviso"
           FechaValida MBFechaCad
           MBFechaCad.SetFocus
        Else
           'Verifica si el año de caducidad es mayor con 2 años al año de Emisión
           If (Aniocad = AniocadAux) Then
              MsgBox "Hola La Fecha de Caducidad no debe sobrepasar dos años, máximo uno", vbInformation, "Aviso"
              FechaValida MBFechaCad
              MBFechaCad.SetFocus
           Else
             If (Aniocad > AniocadAux) Then
                MsgBox "La Fecha de Caducidad no debe sobrepasar dos años, máximo uno", vbInformation, "Aviso"
                FechaValida MBFechaCad
                MBFechaCad.SetFocus
             End If
           End If
        End If
 End If
End Sub

Private Sub MBFechaEmi_LostFocus()
    FechaValida MBFechaEmi
   'Controla que la Fecha de Emisiòn este entre 01/01/2000 en adelante
    If CFechaLong(MBFechaEmi) < CFechaLong("01/01/2000") Then
       MsgBox "La Fecha de Emisión debe ser mayor que 01/01/2000", vbInformation, "Aviso"
       MBFechaEmi = "01/01/2000"
       MBFechaEmi.SetFocus
    End If
    MBFechaRegis = MBFechaEmi
End Sub

Private Sub MBFechaRegis_GotFocus()
    MarcarTexto MBFechaRegis
End Sub

Private Sub MBFechaRegis_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub MBFechaRegis_LostFocus()
   FechaValida MBFechaRegis
   'Controla que la Fecha de Registro este entre 01/01/2000 en adelante
   If CFechaLong(MBFechaRegis) < CFechaLong("01/01/2000") Then
      MsgBox "La Fecha de Registro debe ser mayor que 01/01/2000", vbInformation, "Aviso"
      MBFechaRegis = "01/01/2000"
      MBFechaRegis.SetFocus
   Else
      If MBFechaRegis < MBFechaEmi Then
         MsgBox "La Fecha de Registro debe ser mayor o igual que la Fecha de Emisión", vbInformation, "Aviso"
         MBFechaRegis.SetFocus
      End If
   End If
   FechaValida MBFechaRegis
 ' Carga la Tabla de Porcentaje Iva en el DataCombo
   'Carga_PorcentajeIva (MBFechaRegis)
   Carga_ConceptosRetencion MBFechaRegis
End Sub

Private Sub TxtIvaBienMonIva_GotFocus()
    MarcarTexto TxtIvaBienMonIva
End Sub

Private Sub TxtIvaBienMonIva_LostFocus()
    ' MsgBox CTNumero(TxtIvaBienMonIva, 2)
    TextoValido TxtIvaBienMonIva, True, , 0
End Sub

Private Sub TxtIvaSerMonIva_GotFocus()
    MarcarTexto TxtIvaSerMonIva
End Sub

Private Sub TxtIvaSerMonIva_LostFocus()
Dim Total_IVA_S As Currency
    TextoValido TxtIvaSerMonIva, True, , 0
    'Verifica el Monto Iva Servicios
    Total_IVA_S = CDbl(TxtIvaBienMonIva) + CDbl(TxtIvaSerMonIva)
    If Total_IVA_S > CDbl(TxtMontoIva) Then
       MsgBox "Monto IVA Servicios no puede ser > que Monto IVA", vbInformation, "Aviso de Compras"
       TxtIvaSerMonIva.Text = ""
       TxtIvaSerMonIva.SetFocus
    End If
End Sub

Private Sub TxtIvaSerValRet_GotFocus()
   MarcarTexto TxtIvaSerValRet
End Sub

Private Sub TxtIvaSerValRet_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtMonTitGrat_GotFocus()
   MarcarTexto TxtMonTitGrat
End Sub

Private Sub TxtMonTitGrat_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtMonTitGrat_LostFocus()
   TextoValido TxtMonTitGrat, True, , 2
End Sub

Private Sub TxtMonTitOner_GotFocus()
   MarcarTexto TxtMonTitOner
End Sub

Private Sub TxtMonTitOner_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtMonTitOner_LostFocus()
   TextoValido TxtMonTitOner, True, , 0
End Sub

Private Sub TxtMontoIva_GotFocus()
    MarcarTexto TxtMontoIva
End Sub

Private Sub TxtMontoIva_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtMontoIva_LostFocus()
    TextoValido TxtMontoIva, True, , 0
End Sub

Private Sub TxtNumAutComp_GotFocus()
    MarcarTexto TxtNumAutComp
End Sub

Private Sub TxtNumAutComp_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutComp_LostFocus()
    If Val(TxtNumAutComp) <= 0 Then TxtNumAutComp = "0000000000"
    TxtNumAutComp = Format(Val(Round(TxtNumAutComp)), String(10, "0"))
     
   'Verifico si escogio dcto modificado
   If (cod = 4) Or (cod = 5) Then
      'Selecciona el numero mayor para continuar la secuencia en el
      'campo T_No y A_No
      sSQL = "SELECT TOP 1 * " _
           & "FROM Asiento_Compras " _
           & "WHERE Item = '" & NumEmpresa & "' " _
           & "ORDER BY A_No DESC "
      SelectAdodc AdoAsientoCompras, sSQL
      If AdoAsientoCompras.Recordset.RecordCount > 0 Then Ln_No = AdoAsientoCompras.Recordset.Fields("A_No") + 1
         RatonReloj
         SetAdoAddNew "Asiento_Compras", True
         SetAdoFields "DocModificado", cod
         SetAdoFields "FechaEmiModificado", MBFechaEmiComp
         SetAdoFields "EstabModificado", TxtNumSerieUnoComp
         SetAdoFields "PtoEmiModificado", TxtNumSerieDosComp
         SetAdoFields "SecModificado", CNumSerieTresComp
         SetAdoFields "AutModificado", TxtNumAutComp
         SetAdoFields "MontoTituloOneroso", TxtMonTitOner
         SetAdoFields "MontoTituloGratuito", TxtMonTitGrat
         SetAdoFields "A_No", Maximo_De("Asiento_Compras", "A_No")
         SetAdoFields "T_No", Trans_No
         SetAdoUpdate
      End If
End Sub

Private Sub TxtNumAutor_GotFocus()
     MarcarTexto TxtNumAutor
End Sub

Private Sub TxtNumAutor_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumAutor_LostFocus()
    If Val(TxtNumAutor) <= 0 Then TxtNumAutor = "0000000001"
    TxtNumAutor = Format(Val(Round(TxtNumAutor)), String(10, "0"))
End Sub

Private Sub TxtNumConParPol_GotFocus()
    MarcarTexto TxtNumConParPol
End Sub

Private Sub TxtNumConParPol_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumConParPol_LostFocus()
    TextoValido TxtNumConParPol, True, , 0
    TxtNumConParPol = Format(Val(CCur(TxtNumConParPol)), String(10, "0"))
End Sub

Private Sub TxtNumDosComRet_GotFocus()
   MarcarTexto TxtNumDosComRet
End Sub

Private Sub TxtNumDosComRet_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumDosComRet_LostFocus()
   TextoValido TxtNumDosComRet, True, , 0
   If Val(TxtNumDosComRet) <= 0 Then TxtNumDosComRet = "001"
   TxtNumDosComRet = Format(Val(TxtNumDosComRet), "000")
End Sub

Private Sub TxtNumSerieDos_GotFocus()
    MarcarTexto TxtNumSerieDos
End Sub

Private Sub TxtNumSerieDos_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDos_LostFocus()
    TextoValido TxtNumSerieDos, True, , 0
    If Val(TxtNumSerieDos) <= 0 Then TxtNumSerieDos = "001"
    TxtNumSerieDos = Format(Val(TxtNumSerieDos), "000")
End Sub

Private Sub TxtNumSerieDosComp_GotFocus()
    MarcarTexto TxtNumSerieDosComp
End Sub

Private Sub TxtNumSerieDosComp_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieDosComp_LostFocus()
    TextoValido TxtNumSerieDosComp, True, , 0
    If Val(TxtNumSerieDosComp) <= 0 Then TxtNumSerieDosComp = "001"
       TxtNumSerieDosComp = Format(Val(TxtNumSerieDosComp), "000")
End Sub

Private Sub TxtNumSerietres_GotFocus()
    MarcarTexto TxtNumSerietres
End Sub

Private Sub TxtNumSerietres_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerietres_LostFocus()
    If Val(TxtNumSerietres) <= 0 Then TxtNumSerietres = "000000001"
       TxtNumSerietres = Format(Val(Round(TxtNumSerietres)), "000000000")
End Sub

Private Sub TxtNumSerieUno_GotFocus()
    MarcarTexto TxtNumSerieUno
End Sub

Private Sub TxtNumSerieUno_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieUno_LostFocus()
   TextoValido TxtNumSerieUno, True, , 0
   If Val(TxtNumSerieUno) <= 0 Then TxtNumSerieUno = "001"
      TxtNumSerieUno = Format(Val(TxtNumSerieUno), "000")
End Sub

Private Sub TxtNumSerieUnoComp_GotFocus()
   MarcarTexto TxtNumSerieUnoComp
End Sub

Private Sub TxtNumSerieUnoComp_KeyDown(KeyCode As Integer, Shift As Integer)
   PresionoEnter KeyCode
End Sub

Private Sub TxtNumSerieUnoComp_LostFocus()
    TextoValido TxtNumSerieUnoComp, True, , 0
    If Val(TxtNumSerieUnoComp) <= 0 Then TxtNumSerieUnoComp = "001"
       TxtNumSerieUnoComp = Format(Val(TxtNumSerieUnoComp), "000")
End Sub

Private Sub TxtNumTresComRet_GotFocus()
   MarcarTexto TxtNumTresComRet
End Sub

Private Sub TxtNumTresComRet_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumTresComRet_LostFocus()
   If Val(TxtNumTresComRet) <= 0 Then TxtNumTresComRet = "000000001"
   TxtNumTresComRet = Format(Val(Round(TxtNumTresComRet)), "000000000")
  'Calcula la sumatoria de Monto Iva Bienes, Monto Iva Servicios y Base Imponible
   TxtSumatoria = CTNumero(TxtBaseImpo, 2) + CTNumero(TxtBaseImpoGrav, 2)
  'TxtSumatoria = TxtBaseImpoGrav
   sSQL = "SELECT * " _
        & "FROM Trans_Air " _
        & "WHERE Tipo_Trans = 'C' " _
        & "AND Item = '" & NumEmpresa & "' " _
        & "AND Periodo = '" & Periodo_Contable & "' " _
        & "AND SecRetencion = " & Val(TxtNumTresComRet) & " " _
        & "AND Porcentaje > 0 "
   SelectAdodc AdoAux, sSQL
   If AdoAux.Recordset.RecordCount > 0 Then
      MsgBox "Número de Retención ya existe," & vbCrLf & vbCrLf _
           & "si continua se borrará los datos" & vbCrLf & vbCrLf _
           & "de este número de retención."
   End If
End Sub

Private Sub TxtNumUnoAutComRet_GotFocus()
    MarcarTexto TxtNumUnoAutComRet
End Sub

Private Sub TxtNumUnoAutComRet_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoAutComRet_LostFocus()
    If Val(TxtNumUnoAutComRet) <= 0 Then TxtNumUnoAutComRet = "0000000000"
       TxtNumUnoAutComRet = Format(Val(Round(TxtNumUnoAutComRet)), String(10, "0"))
End Sub

Private Sub TxtNumUnoComRet_GotFocus()
   MarcarTexto TxtNumUnoComRet
End Sub

Private Sub TxtNumUnoComRet_KeyDown(KeyCode As Integer, Shift As Integer)
    PresionoEnter KeyCode
End Sub

Private Sub TxtNumUnoComRet_LostFocus()
   TextoValido TxtNumUnoComRet, True, , 0
   If Val(TxtNumUnoComRet) <= 0 Then TxtNumUnoComRet = "001"
      TxtNumUnoComRet = Format(Val(TxtNumUnoComRet), "000")
End Sub

Public Sub Carga_CreditoTributario()
  ' Carga la Tabla de Catalogos Tributarios al DataCombo
    sSQL = "SELECT (Credito_Tributario & ' - ' & Descripcion) As Sustento,* " _
         & "FROM Tipo_Tributario " _
         & "WHERE Credito_Tributario <> '.' " _
         & "ORDER BY Credito_Tributario "
    SelectDBCombo DCSustento, AdoSustento, sSQL, "Sustento"
End Sub

Public Sub Carga_TipoComprobante(CargaTC As String)
     sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
          & "FROM Tipo_Comprobante " _
          & "WHERE Tipo_Comprobante_Codigo <> 100 " _
          & "ORDER BY Descripcion "
     SelectDBCombo DCDctoModif, AdoTipoComprobante, sSQL, "Descripcion"
    
    'Capturo el codigo del Tipo de Catalogo Tributario
     Cap = CargaTC
            
    'Busco el codigo en la tabla Tipo Comprobante///descripcion
     sSQL = "SELECT CTT.Identificacion,CTT.Tipo_Trans,TC.* " _
          & "FROM Tabla_Tributaria As CTT, Tipo_Comprobante As TC " _
          & "WHERE CTT.Identificacion = '" & CargaTC & "' " _
          & "AND CTT.Tipo_Trans = 'C' "
     If TipoBenef = "R" Then
        sSQL = sSQL & "AND TC.R <> " & Val(adFalse) & " "
     Else
        sSQL = sSQL & "AND TC.C <> " & Val(adFalse) & " "
     End If
     sSQL = sSQL & "AND CTT.Tipo_Comprobante_Codigo = TC.Tipo_Comprobante_Codigo " _
          & "ORDER BY TC.Tipo_Comprobante_Codigo "
     SelectDBCombo DCTipoComprobante, AdoTipoComprobante, sSQL, "Descripcion"
End Sub

Public Sub Captura_TipoComprobante()
   'Captura lo que tiene el Combo de Tipo de Comprobante
    Label15.Caption = "Fechas de " & DCTipoComprobante
    Captc = SinEspaciosIzq(DCTipoComprobante.Text)
    Cap1 = Trim(DCTipoComprobante.Text)
     
    'Busca que sea igual a la Descripcion
    With AdoTipoComprobante.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Descripcion = '" & Cap1 & "' ")
           If Not .EOF Then
              cod = .Fields("Tipo_Comprobante_Codigo")
           Else
              MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
           End If
        End If
    End With
    'MsgBox cod
    If (cod = 4) Or (cod = 5) Then
       FraDctoModificado.Visible = True
       Documento_Modificado
       'Carga en el combo de Documentos Modificados los
       'Tipos de Comprobantes
        sSQL = "SELECT Tipo_Comprobante_Codigo, Descripcion " _
             & "FROM Tipo_Comprobante " _
             & "WHERE Tipo_Comprobante_Codigo <> 100 " _
             & "ORDER BY Descripcion "
        SelectDBCombo DCDctoModif, AdoTipoComprobante, sSQL, "Descripcion"
    Else
        FraDctoModificado.Visible = False
    End If
End Sub

Public Sub Captura_TipoComprobante_DctoModificado()
    CapDcto = Trim(DCDctoModif.Text)
     
    'Busca que sea igual a la Descripcion
    With AdoTipoComprobante.Recordset
        If .RecordCount > 0 Then
           .MoveFirst
           .Find ("Descripcion = '" & CapDcto & "' ")
           If Not .EOF Then
              CapDm = .Fields("Tipo_Comprobante_Codigo")
           Else
              MsgBox "Vuelva a seleccionar", vbInformation, "Aviso"
           End If
        End If
    End With
    If (cod = 4) Or (cod = 5) Then
       FraDctoModificado.Visible = True
      'Verifico si hay documentos modificados de ese Proveedor
       Documento_Modificado
    Else
        FraDctoModificado.Visible = False
    End If
End Sub

Sub Documento_Modificado()
    'Facturas Emitidas del proveedor
     sSQL = "SELECT * " _
          & "FROM Trans_Compras " _
          & "WHERE Item = '" & NumEmpresa & "' " _
          & "AND Periodo = '" & Periodo_Contable & "' " _
          & "AND IdProv = '" & CodigoCliente & "' " _
          & "ORDER BY Secuencial "
     SelectAdodc AdoAux, sSQL
     With AdoAux.Recordset
      If .RecordCount > 0 Then
          Do While Not .EOF
             CNumSerieTresComp.AddItem .Fields("Secuencial")
            .MoveNext
          Loop
      End If
     End With
End Sub

Public Sub Carga_RetencionIvaBienes_Servicios()
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_IVA " _
       & "WHERE Bienes <> " & Val(adFalse) & " " _
       & "ORDER BY Porc "
  SelectDBCombo DCPorcenRetenIvaBien, AdoRetIvaBienes, sSQL, "Porc"
  
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_IVA " _
       & "WHERE Servicios <> " & Val(adFalse) & " " _
       & "ORDER BY Porc "
  SelectDBCombo DCPorcenRetenIvaServ, AdoRetIvaServicios, sSQL, "Porc"
End Sub

Public Sub Carga_ConceptosRetencion(MBFecha As String)
Dim FechaCodAir As String
  FechaCodAir = BuscarFecha(MBFecha)
 'Carga la Tabla de Porcentaje Iva
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE IVA <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc "
  SelectDBCombo DCPorcenIva, AdoPorIva, sSQL, "Porc"
 'Carga los Porcentajes de ICE
  sSQL = "SELECT * " _
       & "FROM Tabla_Por_ICE_IVA " _
       & "WHERE ICE <> " & Val(adFalse) & " " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Porc"
  SelectDBCombo DCPorcenIce, AdoPorIce, sSQL, "Porc"
 'Carga los conceptos de Retencion segun la fecha de Registro
  sSQL = "SELECT (Codigo & ' - ' & Concepto) As Detalle_Conceptos,* " _
       & "FROM Tipo_Concepto_Retencion " _
       & "WHERE Codigo <> '.' " _
       & "AND Fecha_Inicio <= #" & FechaCodAir & "# " _
       & "AND Fecha_Final >= #" & FechaCodAir & "# " _
       & "ORDER BY Codigo "
  SelectDBCombo DCConceptoRet, AdoConceptoRet, sSQL, "Detalle_Conceptos"
  'DCConceptoRet = "329 - Por Otros Servicios (N)"
  'MsgBox sSQL
End Sub

Public Sub Limpiar_Controles()
    ac = 0
    DCRetIBienes.Visible = False
    DCRetISer.Visible = False
    ChRetB.value = False
    ChRetF.value = False
    ChRetS.value = False
    LblNumIdent.Caption = ""
    LblTD.Caption = ""
    OpcNo.value = True
    DCSustento.Text = ""
    DCTipoComprobante.Text = ""
    TxtNumSerieUno.Text = ""
    TxtNumSerieDos.Text = ""
    TxtNumSerietres.Text = ""
    TxtNumAutor.Text = ""
    FechaValida MBFechaEmi
    FechaValida MBFechaRegis
    FechaValida MBFechaCad
    TxtBaseImpo.Text = ""
    TxtBaseImpoGrav.Text = ""
    TxtBaseImpoIce.Text = ""
    DCPorcenIva.Text = ""
    TxtMontoIva.Text = ""
    DCPorcenIce.Text = ""
    TxtMontoIce.Text = ""
    TxtIvaBienMonIva.Text = ""
    DCPorcenRetenIvaBien.Text = ""
    TxtIvaBienValRet.Text = ""
    TxtIvaSerMonIva.Text = ""
    DCPorcenRetenIvaServ.Text = ""
    TxtIvaSerValRet.Text = ""
    TxtNumUnoComRet.Text = ""
    TxtNumDosComRet.Text = ""
    TxtNumTresComRet.Text = ""
    TxtNumUnoAutComRet.Text = ""
    TxtSumatoria.Text = ""
    DCConceptoRet.Text = ""
    TxtBimpConA.Text = ""
    TxtPorRetConA.Text = ""
    TxtValConA.Text = ""
    TxtTotalReten.Text = ""
    TxtNumConParPol.Text = ""
    TxtMonTitOner.Text = ""
    TxtMonTitGrat.Text = ""
    'Limpia la grilla
    'Borra Asiento Air
    sSQL = "DELETE * " _
         & "FROM Asiento_Air " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'C' "
    ConectarAdoExecute sSQL
    
    sSQL = "SELECT * " _
         & "FROM Asiento_Air " _
         & "WHERE codRet <> '.' " _
         & "AND Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " " _
         & "AND Tipo_Trans = 'C' " _
         & "ORDER BY codRet "
    SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL, "Sustento"
End Sub

Public Sub Calculo_Sumatoria()
Dim SumaReten As Currency
  SumaReten = 0
  With AdoAsientoAir.Recordset
   If .RecordCount > 0 Then
      .MoveFirst
       Do While Not .EOF
          SumaReten = SumaReten + .Fields("ValRet")
         .MoveNext
       Loop
   End If
  End With
  TxtTotalReten = Format(SumaReten, "#,##0.00")
End Sub

Public Sub Carga_Datos_Iniciales(MBFecha As MaskEdBox, EsNuevo As Boolean)
    'Encero todo
    ac = 0
    DCPorcenIce = 0
    DCPorcenRetenIvaBien = 0
    DCPorcenRetenIvaServ = 0
    
    CodPorIva = 0
    CodPorIce = "0"
    CodRetBien = 0
    CodRetServ = 0
    
    Limpiar_Controles
    Listar_Air
   'Cargo el No.Autorización de las retenciones
    TxtNumUnoAutComRet = AutorizaRet
   'Carga el Sustento Tributario
    Carga_CreditoTributario
   'Carga en el Data Combo los Clientes con su RUC
    DCTipoComprobante.Text = "Factura"
   'Carga la Tabla de Retencion Iva Bienes y Servicios al DataCombo
    Carga_RetencionIvaBienes_Servicios
    DCPorcenIce.Text = ""
   'Carga la Tabla de Conceptos Retencion al DataCombo
    MBFechaRegis = MBFechaEmi
    Carga_ConceptosRetencion MBFechaEmi
   'Verifico si existen registros caso contrario despliego mensaje
   'Carga los Conceptos de retención en la Fuente al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RF' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRetFuente, AdoRetFuente, sSQL, "Cuentas"
    If AdoRetFuente.Recordset.RecordCount > 0 Then Rf = 1 Else Rf = 0
   'Carga los Conceptos de retención IVA Servicios al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RI' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRetISer, AdoRetIvaSerCC, sSQL, "Cuentas"
    If AdoRetIvaSerCC.Recordset.RecordCount > 0 Then rs = 1 Else rs = 0
    'Carga los Conceptos de retención IVA Bienes al DataCombo
    sSQL = "SELECT (Codigo & ' - ' & Cuenta) As Cuentas  " _
         & "FROM Catalogo_Cuentas " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND Periodo = '" & Periodo_Contable & "' " _
         & "AND TC = 'RI' " _
         & "AND DG = 'D' " _
         & "ORDER BY Codigo "
    SelectDBCombo DCRetIBienes, AdoRetIvaBienesCC, sSQL, "Cuentas"
    If AdoRetIvaBienesCC.Recordset.RecordCount > 0 Then Rb = 1 Else Rb = 0
   'Si es Nuevo ingresa por aqui
    ChRetF.Visible = True
    ChRetF.value = 1
    DCRetFuente.Visible = True
    FrmRetencion.Visible = True
    LblMensaje.Visible = False
    If EsNuevo Then
       'Si todas las variables tienen cero despliego mensaje y no cargo nada
       'No hay cuentas
       If (Rf And rs And Rb) = 0 Then
           ChRetF.Visible = False
           ChRetF.value = 0
           DCRetFuente.Visible = False
           FrmRetencion.Visible = False
           LblMensaje.Visible = True
           Activar_BS
       Else
           ChRetB.SetFocus
       End If
    End If
End Sub

Public Sub Grabacion()
   'Grabo en el Asiento_Compras e implicito Asiento_Air
    If OpcSi.value = True Then ValorP = "S" Else ValorP = "N"
    
    sSQL = "DELETE * " _
         & "FROM Asiento_Compras " _
         & "WHERE Item = '" & NumEmpresa & "' " _
         & "AND CodigoU = '" & CodigoUsuario & "' " _
         & "AND T_No = " & Trans_No & " "
    ConectarAdoExecute sSQL
   'MsgBox Max_IDT("Trans_Compras")
    SetAdoAddNew "Asiento_Compras"
    SetAdoFields "IdProv", CodigoCliente
    SetAdoFields "DevIva", ValorP
    SetAdoFields "CodSustento", Cap
    SetAdoFields "TipoComprobante", cod
    SetAdoFields "Establecimiento", TxtNumSerieUno
    SetAdoFields "PuntoEmision", TxtNumSerieDos
    SetAdoFields "Secuencial", CTNumero(TxtNumSerietres)
    SetAdoFields "Autorizacion", TxtNumAutor
    SetAdoFields "FechaEmision", MBFechaEmi
    SetAdoFields "FechaRegistro", MBFechaRegis
    SetAdoFields "FechaCaducidad", MBFechaCad
    SetAdoFields "BaseImponible", CTNumero(TxtBaseImpo, 2)
    SetAdoFields "BaseImpGrav", CTNumero(TxtBaseImpoGrav, 2)
    SetAdoFields "PorcentajeIva", CodPorIva
    SetAdoFields "MontoIva", CTNumero(TxtMontoIva, 2)
    SetAdoFields "BaseImpIce", CTNumero(TxtBaseImpoIce, 2)
    SetAdoFields "PorcentajeIce", CodPorIce
    SetAdoFields "MontoIce", CTNumero(TxtMontoIce, 2)
    SetAdoFields "Porc_Bienes", DCPorcenRetenIvaBien
    SetAdoFields "MontoIvaBienes", CTNumero(TxtIvaBienMonIva, 2)
    SetAdoFields "PorRetBienes", CodRetBien
    SetAdoFields "ValorRetBienes", CTNumero(TxtIvaBienValRet, 2)
    SetAdoFields "Porc_Servicios", DCPorcenRetenIvaServ
    SetAdoFields "MontoIvaServicios", CTNumero(TxtIvaSerMonIva, 2)
    SetAdoFields "PorRetServicios", CodRetServ
    SetAdoFields "ValorRetServicios", CTNumero(TxtIvaSerValRet, 2)
    
    If (cod = 4) Or (cod = 5) Then
       SetAdoFields "DocModificado", CapDm
       SetAdoFields "FechaEmiModificado", MBFechaEmiComp
       SetAdoFields "EstabModificado", TxtNumSerieUnoComp
       SetAdoFields "PtoEmiModificado", TxtNumSerieDosComp
       SetAdoFields "SecModificado", CNumSerieTresComp
       SetAdoFields "AutModificado", TxtNumAutComp
    Else
       SetAdoFields "DocModificado", "0"
       SetAdoFields "FechaEmiModificado", date
       SetAdoFields "EstabModificado", "000"
       SetAdoFields "PtoEmiModificado", "000"
       SetAdoFields "SecModificado", "0000000"
       SetAdoFields "AutModificado", "0000000000"
    End If
    If TxtNumConParPol = "" Or TxtNumConParPol = "0000000000" Then
       SetAdoFields "ContratoPartidoPolitico", "0000000000"
    Else
       SetAdoFields "ContratoPartidoPolitico", TxtNumConParPol
    End If
    SetAdoFields "MontoTituloOneroso", CTNumero(TxtMonTitOner, 2)
    SetAdoFields "MontoTituloGratuito", CTNumero(TxtMonTitGrat, 2)
   'Verifico si activaron los checks de retenciones
    If ChRetB = 1 Then SetAdoFields "Cta_Bienes", SinEspaciosIzq(DCRetIBienes)
    If ChRetS = 1 Then SetAdoFields "Cta_Servicio", SinEspaciosIzq(DCRetISer)
    SetAdoFields "A_No", 1
    SetAdoFields "T_No", Trans_No
    SetAdoFields "CodigoU", CodigoUsuario
    SetAdoUpdate
    'MsgBox "* ======> " & Trans_No
End Sub

Public Sub Habilita_Controles()
   'Habilito los controles para la modificacion
    SSTCompras.Enabled = True
    CmdGrabar.Enabled = True
    FrmRetencion.Enabled = True
    Label23.Visible = True
End Sub

Public Sub Deshabilita_Controles()
   'Deshabilito los controles para la modificacion
    SSTCompras.Enabled = False
    CmdGrabar.Enabled = False
    FrmRetencion.Enabled = False
    Label23.Visible = False
End Sub

Public Sub Activar_BS()
    TxtIvaBienMonIva.Enabled = True
    DCPorcenRetenIvaBien.Enabled = True
    TxtIvaBienValRet.Enabled = True
    TxtIvaSerMonIva.Enabled = True
    DCPorcenRetenIvaServ.Enabled = True
    TxtIvaSerValRet.Enabled = True
End Sub

Public Sub Listar_Air()
  'Enceramos el espacio de cada usuario para emprezar con una nueva retencion
   sSQL = "DELETE * " _
        & "FROM Asiento_Compras " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
  'Borra Asiento Air
   sSQL = "DELETE * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "AND CodigoU = '" & CodigoUsuario & "' "
   ConectarAdoExecute sSQL
  'Presentamos la malla Asiento Air
  'CodRet,Detalle,BaseImp,Porcentaje,ValRet,EstabRetencion,PtoEmiRetencion,SecRetencion,AutRetencion,FechaEmiRet,Item,CodigoU
   sSQL = "SELECT * " _
        & "FROM Asiento_Air " _
        & "WHERE Item = '" & NumEmpresa & "' " _
        & "AND CodigoU =  '" & CodigoUsuario & "' " _
        & "AND T_No = " & Trans_No & " " _
        & "AND Tipo_Trans = 'C' " _
        & "ORDER BY CodRet "
   SelectDataGrid DGConceptoAir, AdoAsientoAir, sSQL
End Sub

Private Sub TxtPorRetConA_GotFocus()
  MarcarTexto TxtPorRetConA
End Sub

Private Sub TxtPorRetConA_KeyDown(KeyCode As Integer, Shift As Integer)
  PresionoEnter KeyCode
End Sub

Private Sub TxtPorRetConA_LostFocus()
  If OP Then
     TxtValConA = CTNumero(TxtBimpConA, 2) * (CTNumero(TxtPorRetConA, 2) / 100)
     Insertar_DataGrid
  End If
End Sub
