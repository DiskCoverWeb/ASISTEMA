Attribute VB_Name = "Arrays"
Option Explicit

'ubuntu server amd 64
'CREATE TABLE Nombre_Tabla
'(
'    Column1   INT   NOT NULL   INDEX ix1 NONCLUSTERED,
'    Column2 NVARCHAR(4000)
')
'WITH (MEMORY_OPTIMIZED = ON, DURABILITY = SCHEMA_ONLY);

'-----------------------------------
Global BaseEmpresas() As NombreTablas
'-----------------------------------
Global SetD() As Seteos_Documentos
Global FormatoP() As Formatos_Propios
Global LPTs() As Printer
Global TipoC() As Campos_Tabla
Global DatosTabla() As Campos_Tabla
Global CopyGrid() As Campos_Tabla
Global TablaNew() As Crear_Tablas
Global TablaOld() As Crear_Tablas
Global Ancho() As Single
Global Puntos() As Single
Global Campos() As String
Global Maximos()
'Global Fondos() As String
Global SumaTotales() As Variant
Global AnchoDeCampo() As String
'-----------------------------------
Global VString() As String
Global VBTotales() As Byte
Global VLTotales() As Long
Global VITotales() As Integer
Global VCTotales() As Currency
Global VDTotales() As Currency
Global VBTotales1() As Byte
Global VLTotales1() As Long
Global VITotales1() As Integer
Global VCTotales1() As Currency
Global VDTotales1() As Currency
Global DatoGiro() As Datos_Giros
Global GMatriz() As DatoMatriz
'-----------------------------------
Global TotalDia(1 To 7) As Single
Global ImpresorasPapeles() As String
Global LineasLogIn() As String
Global Lista_Archivos() As String
Global Autoizar_Notas(12) As String
Global Formulario As Form
Global CamposRol() As Campos_Rol
Global VCyber_Tiempo() As Cyber_Tiempo
Global Vect_Dec() As Campos_Decimal
Global Fondos_Pantalla() As String
'Global Fondos_Ayuda() As String
Global DBF_Grupo() As String
