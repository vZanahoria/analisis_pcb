Sub CopiarYTransponerColumnasParaTodosLosArchivosVolumen()
    Application.DisplayAlerts = False
    On Error Resume Next

    CarpetaOrigen = "C:\Users\reach\OneDrive\Escritorio\Tec Celaya\Septimo semestre\Introduccion Ciencia de Datos\PruebasAlgoritmos\Pruebas-100-Transformado\"

    ' Abre el libro de destino (el libro al que deseas copiar)
    Dim LibroDestino As Workbook
    Set LibroDestino = Workbooks.Open("C:\Users\reach\OneDrive\Escritorio\Tec Celaya\Septimo semestre\Introduccion Ciencia de Datos\Pruebas\LibroDestino.xlsx")
    If LibroDestino Is Nothing Then
        MsgBox "No se pudo abrir el libro de destino.", vbExclamation
        Exit Sub
    End If

    ' Crea una nueva hoja de destino en el libro de destino
    Set HojaDestino = LibroDestino.Sheets.Add

    ' Itera sobre los archivos de origen en la carpeta
    Archivo = Dir(CarpetaOrigen & "*.xlsx")

    Do While Archivo <> ""
        ' Abre el libro de origen actual
        Set LibroOrigen = Workbooks.Open(CarpetaOrigen & Archivo)

        ' Obtiene el nombre de la primera hoja del libro de origen
        Dim NombrePrimeraHoja As String
        NombrePrimeraHoja = LibroOrigen.Sheets(1).Name

        ' Especifica la hoja de origen
        On Error Resume Next
        Set HojaOrigen = LibroOrigen.Sheets(NombrePrimeraHoja)
        On Error GoTo 0

        ' Verifica que se haya encontrado la hoja de origen
        If HojaOrigen Is Nothing Then
            MsgBox "No se encontró la hoja de origen en el archivo " & Archivo, vbExclamation
            LibroOrigen.Close SaveChanges:=False
            GoTo SiguienteArchivo
        End If

        ' Filtrar el rango para excluir las filas con "GLUE" o "GLUE1" en la columna F
        HojaOrigen.Range("F:F").AutoFilter Field:=1, Criteria1:="<>GLUE", Operator:=xlAnd, Criteria2:="<>GLUE1"

        ' Encuentra la última fila utilizada en la hoja de destino
        UltimaFila = HojaDestino.Cells(HojaDestino.Rows.Count, 2).End(xlUp).Row + 1

        ' Especifica el rango de columnas que deseas copiar en el libro de origen (F, G, M desde la fila 10)
        Set RangoOrigen = Union(HojaOrigen.Range("F10:F" & HojaOrigen.Cells(HojaOrigen.Rows.Count, "F").End(xlUp).Row), _
                                HojaOrigen.Range("G10:G" & HojaOrigen.Cells(HojaOrigen.Rows.Count, "G").End(xlUp).Row), _
                                HojaOrigen.Range("M10:M" & HojaOrigen.Cells(HojaOrigen.Rows.Count, "M").End(xlUp).Row))

        ' Especifica el rango de destino en el libro de destino
        Set RangoDestino = HojaDestino.Range("A" & UltimaFila).Resize(RangoOrigen.Columns.Count, RangoOrigen.Rows.Count)

        ' Escribe el nombre del archivo en la primera columna del rango de destino
        HojaDestino.Cells(HojaDestino.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Left(Archivo, InStrRev(Archivo, ".") - 1)

        ' Copia y transpone el contenido desde el rango de origen al rango de destino
        RangoOrigen.Copy
        RangoDestino.Cells(1, 2).Resize(RangoOrigen.Rows.Count, RangoOrigen.Columns.Count).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

        ' Desactivar el filtro
        HojaOrigen.AutoFilterMode = False

        ' Cierra el libro de origen sin guardar cambios
        LibroOrigen.Close SaveChanges:=False

SiguienteArchivo:
        ' Obtiene el siguiente archivo en la carpeta
        Archivo = Dir
    Loop

    ' Cierra el libro de destino (puedes guardar cambios si es necesario)
    LibroDestino.Close SaveChanges:=True
    Application.DisplayAlerts = True
End Sub
