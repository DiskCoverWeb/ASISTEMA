VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Mayorizar2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "."
   ClientHeight    =   345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   7350
   Begin ComctlLib.ProgressBar ProcBar 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Data DataCuentas 
      Caption         =   "Cuentas"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5565
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Data DataCtas 
      Caption         =   "Ctas"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3885
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Data DataSubCtas 
      Caption         =   "SubCtas"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Data DataTrans 
      Caption         =   "Trans"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Mayorizar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Dim Primero As Boolean
Dim Cod_Cta As String
Dim Num_Comp As Long
  RatonReloj
  Mayorizar2.Caption = "Espere un momento.... Iniciando la mayorizacion."
  sSQL = "SELECT Cta " _
       & "FROM TransaccionesSC " _
       & "WHERE T <> '" & Anulado & "' " _
       & "AND TP <> 'FA' " _
       & "AND TP <> 'N/D' " _
       & "AND TP <> 'N/C' " _
       & "AND Saldo = 0 " _
       & "GROUP BY Cta "
  SelectAdodc DataCuentas, sSQL
  If DataCuentas.Recordset.RecordCount > 0 Then
     DataCuentas.Recordset.MoveLast
     SetProgBar ProcBar, DataCuentas.Recordset.RecordCount
     DataCuentas.Recordset.MoveFirst
     Do While Not DataCuentas.Recordset.EOF
        Codigo4 = DataCuentas.Recordset.Fields("Cta")
       'Seteamos procesos para la mayorizacion.
        Mayorizar2.Caption = "Mayorizando."
        sSQL = "SELECT * FROM TransaccionesSC " _
             & "WHERE T <> '" & Anulado & "' " _
             & "AND TP <> 'FA' " _
             & "AND TP <> 'N/D' " _
             & "AND TP <> 'N/C' " _
             & "AND Cta = '" & Codigo4 & "' " _
             & "ORDER BY Codigo,Factura,Fecha,TP,Numero,Debitos DESC,Creditos,ID "
        SelectAdodc DataSubCtas, sSQL, False
        RatonReloj
        SumaDebe = 0: SumaHaber = 0
        With DataSubCtas.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
             Cod_Cta = .Fields("Codigo")
             SumaCta = 0: Suma_ME = 0
             Mayorizar2.Caption = "Mayorizando la Cuenta: " & Cod_Cta
             Do While Not .EOF
                If Cod_Cta <> .Fields("Codigo") Then
                   Cod_Cta = .Fields("Codigo")
                   SumaCta = 0: Suma_ME = 0
                   Mayorizar2.Caption = "Mayorizando la Cuenta: " & Cod_Cta
                End If
               .Edit
                Debe = Round(.Fields("Debitos"), 2)
                Haber = Round(.Fields("Creditos"), 2)
                Debe_ME = Round(.Fields("Debitos_ME"), 2)
                Haber_ME = Round(.Fields("Creditos_ME"), 2)
                Saldo = Round(.Fields("Saldo"), 2)
                Saldo_ME = Round(.Fields("Saldo_ME"), 2)
                Mayorizar_Saldos .Fields("Cta")
                If ((SumaCta <> Saldo) Or (Suma_ME <> Saldo_ME)) Then
                  .Fields("Saldo") = SumaCta
                  .Fields("Saldo_ME") = Suma_ME
                  .Update
                End If
               .MoveNext
             Loop
         End If
        End With
        IncProgBar ProcBar
        DataCuentas.Recordset.MoveNext
     Loop
  End If
  RatonNormal
  sSQL = "SELECT Cta " _
       & "FROM Transacciones " _
       & "WHERE T <> '" & Anulado & "' " _
       & "AND TP <> 'FA' " _
       & "AND TP <> 'N/D' " _
       & "AND TP <> 'N/C' " _
       & "AND Saldo = 0 " _
       & "GROUP BY Cta "
  SelectAdodc DataCuentas, sSQL
  If DataCuentas.Recordset.RecordCount > 0 Then
     DataCuentas.Recordset.MoveLast
     SetProgBar ProcBar, DataCuentas.Recordset.RecordCount
     DataCuentas.Recordset.MoveFirst
     Do While Not DataCuentas.Recordset.EOF
        Codigo4 = DataCuentas.Recordset.Fields("Cta")
        sSQL = "SELECT * FROM Transacciones " _
             & "WHERE T <> '" & Anulado & "' " _
             & "AND TP <> 'FA' " _
             & "AND TP <> 'N/D' " _
             & "AND TP <> 'N/C' " _
             & "AND Cta = '" & Codigo4 & "' " _
             & "ORDER BY Cta,Fecha,TP,Numero,Debe DESC,Haber,ID "
        SelectAdodc DataTrans, sSQL, False
        RatonReloj
        SumaDebe = 0: SumaHaber = 0
        With DataTrans.Recordset
         If .RecordCount > 0 Then
            .MoveFirst
             Cod_Cta = .Fields("Cta")
             SumaCta = 0: Suma_ME = 0
             Mayorizar2.Caption = "Mayorizando la Cuenta: " & Cod_Cta
             Do While Not .EOF
                If Cod_Cta <> .Fields("Cta") Then
                   Cod_Cta = .Fields("Cta")
                   SumaCta = 0: Suma_ME = 0
                   Mayorizar2.Caption = "Mayorizando la Cuenta: " & Cod_Cta
                End If
               .Edit
                Debe = Round(.Fields("Debe"), 2)
                Haber = Round(.Fields("Haber"), 2)
                Debe_ME = Round(.Fields("Debe_ME"), 2)
                Haber_ME = Round(.Fields("Haber_ME"), 2)
                Saldo = Round(.Fields("Saldo"), 2)
                Saldo_ME = Round(.Fields("Saldo_ME"), 2)
                Mayorizar_Saldos .Fields("Cta")
                If ((SumaCta <> Saldo) Or (Suma_ME <> Saldo_ME)) Then
                  .Fields("Saldo") = SumaCta
                  .Fields("Saldo_ME") = Suma_ME
                  .Update
                End If
                SumaDebe = SumaDebe + Debe
                SumaHaber = SumaHaber + Haber
               .MoveNext
             Loop
         End If
        End With
        IncProgBar ProcBar
        DataCuentas.Recordset.MoveNext
     Loop
  End If
  RatonNormal
  If Round(SumaDebe - SumaHaber, 2) <> 0 Then
     Cadena = "Warning: Verifique las Transacciones," & vbCrLf & vbCrLf
     Cadena = Cadena & "                No Cuadran por: " & Format(Abs(SumaDebe - SumaHaber), "#,##0.00") & "."
     MsgBox Cadena
  End If
  Unload Mayorizar2
End Sub

Private Sub Form_Load()
  CentrarForm Mayorizar2
  DataCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataTrans.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataCuentas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  DataSubCtas.DatabaseName = RutaEmpresa & "\CONTABIL.MDB"
  SetProgBar ProcBar, 100
End Sub

