Attribute VB_Name = "modIDs"
Public Sub GenerarID(ByVal ws As Worksheet, ByVal Target As Range, ByVal nombreTabla As String, ByVal nombreColumnaID As String)
    On Error GoTo ManejoError
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim celdaID As Range
    Dim maxID As Long
    Dim tblParametros As ListObject
    Dim filaParametro As ListRow
    Dim encontrado As Boolean
    
    Set tbl = ws.ListObjects(nombreTabla)
    Set tblParametros = ws.Parent.Worksheets("PARAMETROS").ListObjects("tblParametrosID")
    
    If Intersect(Target, tbl.DataBodyRange) Is Nothing Then Exit Sub

    Set fila = tbl.ListRows(Target.Row - tbl.DataBodyRange.Row + 1)
    Set celdaID = fila.Range.Columns(tbl.ListColumns(nombreColumnaID).Index)
    
    If celdaID.Value = "" Then
        For Each filaParametro In tblParametros.ListRows
            encontrado = False
            If filaParametro.Range.Columns(tblParametros.ListColumns("CLAVE").Index).Value = nombreColumnaID Then
                encontrado = True
                If Not IsNumeric(filaParametro.Range.Columns(tblParametros.ListColumns("ULTIMO_ID").Index).Value) Then
                    MsgBox "ULTIMO_ID de " & nombreColumnaID & " no contiene un valor numérico válido.", vbExclamation
                    Exit Sub
                End If
                maxID = filaParametro.Range.Columns(tblParametros.ListColumns("ULTIMO_ID").Index).Value
                If maxID < 0 Then
                    MsgBox "ULTIMO_ID de " & nombreColumnaID & " no puede ser negativo.", vbExclamation
                    Exit Sub
                End If
                Application.EnableEvents = False
                filaParametro.Range.Columns(tblParametros.ListColumns("ULTIMO_ID").Index).Value = maxID + 1
                celdaID.Value = maxID + 1
                Application.EnableEvents = True
            Exit For
            End If
        Next filaParametro
        If encontrado = False Then
            MsgBox "No se encontró la clave " & nombreColumnaID & " en tblParametrosID.", vbExclamation
        End If
    End If
        'If WorksheetFunction.Count(tbl.ListColumns(nombreColumnaID).DataBodyRange) = 0 Then
         '   maxID = 0
        'Else
         '   maxID = WorksheetFunction.Max(tbl.ListColumns(nombreColumnaID).DataBodyRange)
        'End If
        
        'Application.EnableEvents = False
        'celdaID.Value = maxID + 1
        'Application.EnableEvents = True
    'End If
    Exit Sub
    
ManejoError:
    Application.EnableEvents = True
    MsgBox "Ocurrió un error al generar el ID: " & Err.Description, vbExclamation
End Sub
