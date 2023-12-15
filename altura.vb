Sub CopiarYTransponerColumnasParaTodosLosArchivosAltura()
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

    ' Especifica la hoja de destino
    On Error Resume Next
    Set HojaDestino = LibroDestino.Sheets("Hoja1")
    On Error GoTo 0
    If HojaDestino Is Nothing Then
        MsgBox "No se encontró la hoja de destino.", vbExclamation
        LibroDestino.Close SaveChanges:=False
        Exit Sub
    End If

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

        ' Especifica el rango de columnas que deseas copiar en el libro de origen (F:L desde la fila 10)
        Set RangoOrigen = HojaOrigen.Range("F10:I" & HojaOrigen.Cells(HojaOrigen.Rows.Count, "F").End(xlUp).Row)

        ' Filtrar el rango para excluir las filas con "GLUE" o "GLUE1" en la columna F
        RangoOrigen.AutoFilter Field:=1, Criteria1:="<>GLUE", Operator:=xlAnd, Criteria2:="<>GLUE1"

        ' Copia y transpone el contenido visible desde el rango de origen al rango de destino
        RangoOrigen.SpecialCells(xlCellTypeVisible).Copy
        HojaDestino.Cells(HojaDestino.Rows.Count, 2).End(xlUp).Offset(1, 0).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

        ' Desactivar el filtro
        HojaOrigen.AutoFilterMode = False

        ' Escribe el nombre del archivo en la primera columna del rango de destino
        HojaDestino.Cells(HojaDestino.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Left(Archivo, InStrRev(Archivo, ".") - 1)
        HojaDestino.Cells(HojaDestino.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Left(Archivo, InStrRev(Archivo, ".") - 1)
        HojaDestino.Cells(HojaDestino.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Left(Archivo, InStrRev(Archivo, ".") - 1)
        HojaDestino.Cells(HojaDestino.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Left(Archivo, InStrRev(Archivo, ".") - 1)

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