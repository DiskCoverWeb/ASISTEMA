Attribute VB_Name = "SubEmpresa"
Option Explicit

Public Function Datos_De_Empresa(BuscarCampo As String) As Variant
Dim RegEmp As ADODB.Recordset
Dim DatosSelect As String
Dim IndDato As Variant
Dim JCamp As Integer
    RatonReloj
    DatosSelect = "SELECT * " _
                & "FROM Empresas " _
                & "WHERE Item = '" & NumEmpresa & "' "
    DatosSelect = CompilarSQL(DatosSelect)
    Set RegEmp = New ADODB.Recordset
    RegEmp.CursorType = adOpenStatic
    RegEmp.CursorLocation = adUseClient
    RegEmp.Open DatosSelect, AdoStrCnn, , , adCmdText
    With RegEmp
     If .RecordCount > 0 Then
         IndDato = ""
         For JCamp = 0 To .Fields.Count - 1
          If BuscarCampo = .Fields(JCamp).Name Then IndDato = .Fields(BuscarCampo)
         Next JCamp
     Else
         IndDato = ""
     End If
    End With
    RegEmp.Close
    RatonNormal
    Datos_De_Empresa = IndDato
End Function
