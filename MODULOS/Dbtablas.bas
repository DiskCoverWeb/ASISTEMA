Attribute VB_Name = "CrearTablas"
'====================================================
'|Tipo |Ancho|Prec_|Descripción
'|Campo|Campo|isión|y Ejemplo
'====================================================
'|  C  |  n  |  -  |Campo de caracteres de ancho n
'|  d  |  -  |  -  |Date
'|  T  |  -  |  -  |DateTime
'|  N  |  n  |  d  |Campo numérico de ancho n con d decimales
'|  F  |  n  |  d  |Campo numérico flotante de ancho n con d decimales
'|  I  |  -  |  -  |Integer
'|  B  |  -  |  d  |Currency
'|  Y  |  -  |  -  |Currency
'|  L  |  -  |  -  |Logical
'|  M  |  -  |  -  |Memo
'|  G  |  -  |  -  |General
'=========================================================
'Si existe una tabla con el nombre que se le asigna, entonces se elimina,
'caso contrarario se crea una tabla con nuevos campos

''''Entrada :  DBs1,NombTbl
''''Salida : ExisteSN
''''Verifica si existe o no el nombre de la tabla que se recibe en la base de datos
'''Public Function NoExisteTabla(DBs1() As NombreTablas, NombTbl As String) As Boolean
'''Dim ExisteSN As Boolean
'''    ExisteSN = True
'''    For J = 0 To DBs1(0).Cantidad
'''     If DBs1(J).Nombre = NombTbl Then
'''        ExisteSN = False
'''       'MsgBox NombTbl & " =-> " & DBs1(J).Nombre
'''     End If
'''    Next
'''    NoExisteTabla = ExisteSN
'''End Function
'''
'''Public Sub CTAsientoContable()
'''Dim CTabla As String
'''Dim NombreTabla As String
'''  NombreTabla = "Asiento_C_" & CodigoUsuario
'''  CTabla = "CREATE TABLE " & NombreTabla & " (" _
'''         & "C BIT," _
'''         & "FECHA DATETIME," _
'''         & "BENEFICIARIO TEXT(35)," _
'''         & "TP TEXT(3)," _
'''         & "NUMERO LONG," _
'''         & "CHEQ_DEP TEXT(8)," _
'''         & "DEBE Currency ," _
'''         & "HABER Currency ," _
'''         & "ME BIT," _
'''         & "Item BYTE," _
'''         & "T_No BYTE" _
'''         & "); "
'''  If NoExisteTabla(BaseEmpresas, NombreTabla) Then Conectar_Ado_Execute CTabla
'''End Sub
'''
'''Public Sub Tablas_de_Base(Dta As Data, NombreBase() As NombreTablas)
'''Dim TD As Database
'''  J = 0
'''  For Each TD In Dta.Database.TableDefs
'''      If UCase(MidStrg(TD.Name, 1, 4)) <> "MSYS" Then J = J + 1
'''  Next
'''  ReDim NombreBase(J) As NombreTablas
'''  I = 0
'''  For Each TD In Dta.Database.TableDefs
'''   If UCase(MidStrg(TD.Name, 1, 4)) <> "MSYS" Then
'''      NombreBase(I).Nombre = TD.Name
'''      NombreBase(I).Cantidad = J - 1
'''      I = I + 1
'''   End If
'''  Next
'''End Sub
