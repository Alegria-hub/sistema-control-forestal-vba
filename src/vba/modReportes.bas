Attribute VB_Name = "modReportes"
Option Explicit


Public Sub PrepararReporteNuevo()

    Dim wsReporte As Worksheet
    Dim ultimaFila As Long

    Set wsReporte = ThisWorkbook.Worksheets("REPORTES")

    ' Identificar hasta dónde llega el reporte anterior
    ultimaFila = wsReporte.Cells(wsReporte.Rows.Count, "A").End(xlUp).Row

    If ultimaFila < 12 Then ultimaFila = 12

    ' Eliminar las tablas anteriores
    wsReporte.Range("A10:F" & ultimaFila).UnMerge
    wsReporte.Range("A10:F" & ultimaFila).Clear

    ' Recuperar nuestra plantilla maestra
    wsReporte.Range("H10:M12").Copy _
        Destination:=wsReporte.Range("A10")

    ' Limpiar fecha anterior
    wsReporte.Range("C7:D7").ClearContents

    Application.CutCopyMode = False

End Sub


Public Function InsertarEncabezadoEmpresa(ByVal nombreEmpresa As String) As Long

    Dim wsReporte As Worksheet
    Dim ultimaCelda As Range
    Dim filaInicio As Long

    Set wsReporte = ThisWorkbook.Worksheets("REPORTES")

    ' Comprobar si utilizaremos el primer bloque
    If wsReporte.Range("A10").Value = "NOMBRE DE LA EMPRESA" Then

        filaInicio = 10

    Else

        ' Buscar dónde termina el último bloque
        Set ultimaCelda = wsReporte.Range("A10:F" & _
            wsReporte.Rows.Count).Find( _
            What:="*", _
            LookIn:=xlValues, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious)

        filaInicio = ultimaCelda.Row + 2
        
        wsReporte.Rows(filaInicio - 1).RowHeight = 12


        ' Copiar el diseño de la plantilla maestra
        wsReporte.Range("H10:M12").Copy _
            Destination:=wsReporte.Cells(filaInicio, 1)

        wsReporte.Rows(filaInicio).RowHeight = 28
        wsReporte.Rows(filaInicio + 2).RowHeight = 55

    End If
    
    ' Altura uniforme del espacio entre encabezado y tabla
    wsReporte.Rows(filaInicio + 1).RowHeight = 12


    ' Escribir el nombre de la empresa
    wsReporte.Cells(filaInicio, 1).Value = nombreEmpresa

    Application.CutCopyMode = False
    
    InsertarEncabezadoEmpresa = filaInicio
End Function


Public Sub InsertarDetalleEmpresa( _
    ByVal idEmpresa As String, _
    ByVal fechaBalance As Date, _
    ByVal filaInicio As Long)

    Dim wsReporte As Worksheet
    Dim tblBalance As ListObject
    Dim tblEspecies As ListObject
    Dim filaBalance As ListRow
    Dim filaEspecie As ListRow
    Dim filaDetalle As Long
    
    Dim idRegistro As String
    Dim fechaRegistro As Variant
    Dim idEspecie As String

    Set wsReporte = ThisWorkbook.Worksheets("REPORTES")

    Set tblBalance = ThisWorkbook.Worksheets("BALANCE") _
        .ListObjects("tblBalance")

    Set tblEspecies = ThisWorkbook.Worksheets("CATALOGOS") _
        .ListObjects("tblEspecies")

    filaDetalle = filaInicio + 3
    
    For Each filaBalance In tblBalance.ListRows
    
        idRegistro = CStr(filaBalance.Range.Cells(1, _
            tblBalance.ListColumns("ID_EMPRESA").Index).Value)
    
        fechaRegistro = filaBalance.Range.Cells(1, _
            tblBalance.ListColumns("FECHA_BALANCE").Index).Value
    
        If IsDate(fechaRegistro) Then
    
            If idRegistro = idEmpresa And _
                DateValue(CDate(fechaRegistro)) = DateValue(fechaBalance) Then
    
                'Debug.Print "ID_ESPECIE: "; filaBalance.Range.Cells(1, _
                 '   tblBalance.ListColumns("ID_ESPECIE").Index).Value
                
idEspecie = CStr(filaBalance.Range.Cells(1, _
    tblBalance.ListColumns("ID_ESPECIE").Index).Value)

        For Each filaEspecie In tblEspecies.ListRows
        
            If CStr(filaEspecie.Range.Cells(1, _
                tblEspecies.ListColumns("ID_ESPECIE").Index).Value) = idEspecie Then
        
                'Debug.Print "ESPECIE: "; filaEspecie.Range.Cells(1, _
                 '   tblEspecies.ListColumns("ESPECIE").Index).Value
                
                wsReporte.Cells(filaDetalle, "A").Value = _
                    filaEspecie.Range.Cells(1, _
                    tblEspecies.ListColumns("ESPECIE").Index).Value
                    
                ' Inventario inspeccionado
                wsReporte.Cells(filaDetalle, "B").Value = _
                    filaBalance.Range.Cells(1, _
                    tblBalance.ListColumns("M3_BASE").Index).Value
                
                ' Entradas
                wsReporte.Cells(filaDetalle, "C").Value = _
                    filaBalance.Range.Cells(1, _
                    tblBalance.ListColumns("ENTRADAS_M3").Index).Value
                
                ' Salidas
                wsReporte.Cells(filaDetalle, "D").Value = _
                    filaBalance.Range.Cells(1, _
                    tblBalance.ListColumns("SALIDAS_M3").Index).Value
                
                ' Inventario actual
                wsReporte.Cells(filaDetalle, "E").Value = _
                    filaBalance.Range.Cells(1, _
                    tblBalance.ListColumns("INVENTARIO_FISICO_M3").Index).Value
                
                ' Balance
                wsReporte.Cells(filaDetalle, "F").Value = _
                    filaBalance.Range.Cells(1, _
                    tblBalance.ListColumns("BALANCE_M3").Index).Value
                
                ' Mostrar tres decimales
                wsReporte.Range("B" & filaDetalle & ":F" & filaDetalle).NumberFormat = "0.000"
                
                
                ' Aplicar formato a la fila de la especie
                With wsReporte.Range( _
                    wsReporte.Cells(filaDetalle, 1), _
                    wsReporte.Cells(filaDetalle, 6))
                
                    .Interior.Color = RGB(244, 247, 252)
                
                    .Font.Name = "Aptos"
                    .Font.Size = 10
                    .Font.Color = RGB(51, 51, 51)
                
                    .VerticalAlignment = xlCenter
                
                    .Borders.LineStyle = xlContinuous
                    .Borders.Color = RGB(218, 224, 235)
                    .Borders.Weight = xlThin
                
                End With
                
                ' Ajustar presentación de la fila
                wsReporte.Rows(filaDetalle).RowHeight = 23
                
                wsReporte.Cells(filaDetalle, 1).HorizontalAlignment = xlLeft
                
                wsReporte.Range( _
                    wsReporte.Cells(filaDetalle, 2), _
                    wsReporte.Cells(filaDetalle, 6)).HorizontalAlignment = xlRight

                
                filaDetalle = filaDetalle + 1

                Exit For
        
            End If
        
        Next filaEspecie

            End If
    
        End If
    
    Next filaBalance


End Sub


Public Sub ConfigurarImpresionReporte()
    Dim wsReporte As Worksheet
    Dim ultimaFila As Long

    Set wsReporte = ThisWorkbook.Worksheets("REPORTES")

    ' Encontrar la ultima fila del reporte
    ultimaFila = wsReporte.Cells( _
        wsReporte.Rows.Count, "A").End(xlUp).Row

    If ultimaFila < 12 Then ultimaFila = 12

    ' Configurar la pagina
    With wsReporte.PageSetup

        .PrintArea = "$A$1:$F$" & ultimaFila

        .Orientation = xlPortrait
        .PaperSize = xlPaperLetter

        ' Ajustar al ancho de una pagina
        .Zoom = 130
        .FitToPagesWide = 1
        .FitToPagesTall = False

        ' Centrar horizontalmente
        .CenterHorizontally = True
        .CenterVertically = False

        ' Ajustar margenes
        .LeftMargin = Application.InchesToPoints(0.1)
        .RightMargin = Application.InchesToPoints(0.1)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)

        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)

        ' Ocultar las lineas de cuadricula
        .PrintGridlines = False

    End With

    MsgBox "Configuracion de impresion aplicada.", vbInformation
End Sub


Public Sub ExportarReportePDF()

    Dim wsReporte As Worksheet
    Dim ultimaFila As Long
    Dim rutaPDF As Variant
    Dim nombreArchivo As String

    Set wsReporte = ThisWorkbook.Worksheets("REPORTES")

    ' Detectar la ultima fila del reporte
    ultimaFila = wsReporte.Cells( _
        wsReporte.Rows.Count, "A").End(xlUp).Row

    If ultimaFila < 13 Then
        MsgBox "Agrega al menos una empresa con registros antes de exportar.", _
               vbExclamation, "Reporte vacio"
        Exit Sub
    End If

    ' Actualizar el area de impresion
    wsReporte.PageSetup.PrintArea = "$A$1:$F$" & ultimaFila

    ' Proponer un nombre para el archivo
    nombreArchivo = "Reporte_Balance_Forestal_" & _
                    Format(wsReporte.Range("C7").Value, "dd-mm-yyyy") & ".pdf"

    ' Elegir la ubicacion donde guardar el PDF
    rutaPDF = Application.GetSaveAsFilename( _
        InitialFileName:=nombreArchivo, _
        FileFilter:="Archivos PDF (*.pdf), *.pdf", _
        Title:="Guardar reporte de balance forestal")

    If rutaPDF = False Then Exit Sub

    ' Asegurar la extension PDF
    If LCase$(Right$(CStr(rutaPDF), 4)) <> ".pdf" Then
        rutaPDF = CStr(rutaPDF) & ".pdf"
    End If

    ' Exportar respetando la configuracion de impresion
    wsReporte.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=CStr(rutaPDF), _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Reporte exportado correctamente a PDF.", _
           vbInformation, "Exportacion completada"

End Sub



