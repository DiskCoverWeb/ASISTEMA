VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClaseExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ClaseExcel:
 
Option Explicit
 
'devuelve un objeto Recordset con los datos de la hoja
Public Function Leer_Excel(ByVal PathXls As String) As ADODB.Recordset
Dim obj_Excel     As Object
Dim obj_Workbook  As Object
Dim obj_Worksheet As Object
 
Dim rs As ADODB.Recordset
   
Dim hoja As String
Dim cs As String
Dim sSheetName As String
 
   Set rs = New ADODB.Recordset
   RatonReloj
   FEsperar.Show
   Imagen_Esperar "Importando el Archivo: " & vbCrLf & PathXls

   rs.CursorLocation = adUseClient
   rs.CursorType = adOpenKeyset
   rs.LockType = adLockBatchOptimistic
 
   cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
   ' -- crea rnueva instancia de Excel
   Set obj_Excel = CreateObject("Excel.Application")
 
   'obj_Excel.Visible = True
 
   ' -- Abrir el libro
   Set obj_Workbook = obj_Excel.Workbooks.open(PathXls)
   ' -- referencia la Hoja, por defecto la hoja activa
   If sSheetName = vbNullString Then
      Set obj_Worksheet = obj_Workbook.ActiveSheet
      hoja = obj_Workbook.ActiveSheet.Name
   Else
      'Set obj_Worksheet = obj_Workbook.Sheets(sSheetName)
      hoja = obj_Workbook.Sheets(sSheetName)
   End If
 
   hoja = "[" & hoja & "$" & "]"
   rs.open "SELECT * FROM " & hoja, cs
 
   Set Leer_Excel = rs
   Set rs = Nothing
 
   ' -- Cerrar libro
   obj_Workbook.Close
   ' -- Cerrar Excel
   obj_Excel.Quit
   Set obj_Workbook = Nothing
   Set obj_Excel = Nothing
   Set obj_Worksheet = Nothing
   
   RatonNormal
   Unload FEsperar
End Function

Public Sub Exportar_Excel(consulta As ADODB.Recordset, Optional nombreHoja As String)
Dim APIExcel As Object
Dim AddLibro As Object
Dim AddHoja  As Object

Dim I As Long
Dim columnas As Long
Dim filas As Long

'Creamos objeto excel y nuevo libro y no mostramos el archivo
Set APIExcel = CreateObject("Excel.Application")
Set AddLibro = APIExcel.Workbooks.Add
APIExcel.Visible = False
'Añadimos hoja al libro nuevo y nombramos pestaña
Set AddHoja = AddLibro.Worksheets(1)

'Damos nombre a la hoja con la que vamos a exportar los datos
If Len(nombreHoja) > 0 Then AddHoja.Name = Left(nombreHoja, 60) Else AddHoja.Name = "DiskCover System"
'Traemos los datos de cabecera de la tabla Access y los pegamos en la hoja excel

columnas = consulta.Fields.Count
filas = consulta.RecordCount

'Generamos encabezado con colores
With APIExcel.Range(APIExcel.cells(2, 1), APIExcel.cells(2, columnas))
.Font.Bold = True
.Interior.color = RGB(168, 168, 0)
.HorizontalAlignment = 3
.VerticalAlignment = 2
.EntireRow.RowHeight = 20
End With

With APIExcel.Range(APIExcel.cells(2, 1), APIExcel.cells(filas + 2, columnas)).Borders
.LineStyle = 1
.Weight = 1
.ColorIndex = 5
End With

For I = 0 To columnas - 1
APIExcel.cells(2, I + 1) = consulta.Fields(I).Name
Next I

'Pegamos los datos de la tabla en la nueva hoja
consulta.MoveFirst
AddHoja.Range("A3").CopyFromRecordset consulta
'Damos formato a las columnas, ajustando contenidos
With APIExcel.ActiveSheet.cells
.Select
.EntireColumn.AutoFit
.Range("A1").Select
End With
'Mostramos la hoja
APIExcel.Visible = True
End Sub

