VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form SeteoEmp 
   Caption         =   "Setear datos de empresa"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   6285
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1995
      Width           =   1905
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   1275
      Left            =   105
      OleObjectBlob   =   "SeteoEmp.frx":0000
      TabIndex        =   0
      Top             =   105
      Width           =   6105
   End
End
Attribute VB_Name = "SeteoEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

