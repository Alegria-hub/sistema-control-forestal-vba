VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReporteBalance 
   Caption         =   "GENERADOR DE REPORTES FORESTALES"
   ClientHeight    =   5160
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmReporteBalance.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmReporteBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboEmpresa_Change()
    If cboEmpresa.ListIndex = -1 Then Exit Sub

    Debug.Print "Empresa: " & cboEmpresa.List(cboEmpresa.ListIndex, 0)
    Debug.Print "ID_EMPRESA: " & cboEmpresa.List(cboEmpresa.ListIndex, 1)

End Sub

Private Sub cboFechaBalance_Change()
    Dim fechaSeleccionada As String
    Dim tblBalance As ListObject
    Dim fila As ListRow
    Dim fechaRegistro As Variant
    Dim empresasUnicas As Object
    Dim idEmpresa As String
    
    Dim tblEmpresas As ListObject
    Dim filaEmpresa As ListRow
    
    cboEmpresa.Clear
    cboEmpresa.Enabled = False
    lstEmpresasAgregadas.Clear
    PrepararReporteNuevo

    If cboFechaBalance.ListIndex = -1 Then Exit Sub

    fechaSeleccionada = cboFechaBalance.Value
    
    ThisWorkbook.Worksheets("REPORTES").Range("C7").Value = _
        DateSerial( _
            CInt(Right(fechaSeleccionada, 4)), _
            CInt(Mid(fechaSeleccionada, 4, 2)), _
            CInt(Left(fechaSeleccionada, 2)) _
        )


    Debug.Print "Fecha seleccionada: [" & fechaSeleccionada & "]"

    Set tblBalance = ThisWorkbook.Worksheets("BALANCE").ListObjects("tblBalance")
    Set tblEmpresas = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEmpresas")
    Set empresasUnicas = CreateObject("Scripting.Dictionary")
    
    For Each fila In tblBalance.ListRows
    
        fechaRegistro = fila.Range.Cells(1, _
            tblBalance.ListColumns("FECHA_BALANCE").Index).Value
    
        If IsDate(fechaRegistro) Then
    
            If Format(CDate(fechaRegistro), "dd/mm/yyyy") = fechaSeleccionada Then
    
                'Debug.Print "ID_EMPRESA: "; fila.Range.Cells(1, _
                 '   tblBalance.ListColumns("ID_EMPRESA").Index).Value
                 
                idEmpresa = CStr(fila.Range.Cells(1, _
                    tblBalance.ListColumns("ID_EMPRESA").Index).Value)
                
                If Not empresasUnicas.Exists(idEmpresa) Then
                
                    empresasUnicas.Add idEmpresa, True
                
                    'Debug.Print "ID_EMPRESA: "; idEmpresa
                    
                     For Each filaEmpresa In tblEmpresas.ListRows

                        If CStr(filaEmpresa.Range.Cells(1, _
                            tblEmpresas.ListColumns("ID_EMPRESA").Index).Value) = idEmpresa Then
                
                            'Debug.Print "EMPRESA: "; filaEmpresa.Range.Cells(1, _
                             '   tblEmpresas.ListColumns("EMPRESA").Index).Value
                             
                            cboEmpresa.AddItem filaEmpresa.Range.Cells(1, _
                                tblEmpresas.ListColumns("EMPRESA").Index).Value
                            
                            cboEmpresa.List(cboEmpresa.ListCount - 1, 1) = idEmpresa
                            

                
                            Exit For
                
                        End If
                
                    Next filaEmpresa
        
                End If
            
            End If
    
        End If
    
    Next fila
    cboEmpresa.Enabled = (cboEmpresa.ListCount > 0)
End Sub

Private Sub cmdAgregarEmpresa_Click()
    Dim i As Long
    Dim filaInicio As Long

    If cboEmpresa.ListIndex = -1 Then
        MsgBox "Selectiona una empresa.", vbExclamation
        Exit Sub
    End If

    For i = 0 To lstEmpresasAgregadas.ListCount - 1

        If lstEmpresasAgregadas.List(i) = cboEmpresa.Value Then
            MsgBox "Esta empresa ya fue agregada.", vbExclamation
            Exit Sub
        End If

    Next i

    lstEmpresasAgregadas.AddItem cboEmpresa.Value
    
    lstEmpresasAgregadas.List(lstEmpresasAgregadas.ListCount - 1, 1) = _
        cboEmpresa.List(cboEmpresa.ListIndex, 1)
        
    
    'InsertarEncabezadoEmpresa cboEmpresa.Value
    
    filaInicio = InsertarEncabezadoEmpresa(cboEmpresa.Value)
    
    InsertarDetalleEmpresa _
        CStr(cboEmpresa.List(cboEmpresa.ListIndex, 1)), _
        CDate(ThisWorkbook.Worksheets("REPORTES").Range("C7").Value), _
        filaInicio

    
    Debug.Print "Encabezado creado en fila: "; filaInicio

        
    
    'Debug.Print "Empresa agregada test: " & _
     '   lstEmpresasAgregadas.List(lstEmpresasAgregadas.ListCount - 1, 0)
    
    'Debug.Print "ID_EMPRESA test: " & _
     '   lstEmpresasAgregadas.List(lstEmpresasAgregadas.ListCount - 1, 1)


End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdLimpiarReporte_Click()
    Dim respuesta As VbMsgBoxResult

    respuesta = MsgBox( _
        "¿Deseas iniciar un nuevo reporte? Se limpiará el reporte actual.", _
        vbYesNo + vbQuestion, _
        "Nuevo reporte")

    If respuesta = vbNo Then Exit Sub

    PrepararReporteNuevo

    cboFechaBalance.ListIndex = -1
    cboEmpresa.Clear
    cboEmpresa.Enabled = False
    lstEmpresasAgregadas.Clear

    MsgBox "El reporte está listo para comenzar.", vbInformation
End Sub

Private Sub PDFcmdGenerar_Click()
        ExportarReportePDF
End Sub

Private Sub UserForm_Initialize()
    Dim tblBalance As ListObject
    Dim fila As ListRow
    Dim fechas As Object
    Dim fecha As Variant
    Dim claveFecha As String
    
    cboEmpresa.Enabled = False
    
    cboEmpresa.ColumnCount = 2
    cboEmpresa.ColumnWidths = "250 pt;0 pt"
    
    
    lstEmpresasAgregadas.ColumnCount = 2
    lstEmpresasAgregadas.ColumnWidths = "250 pt;0 pt"


    Set tblBalance = ThisWorkbook.Worksheets("BALANCE").ListObjects("tblBalance")

    Set fechas = CreateObject("Scripting.Dictionary")

    cboFechaBalance.Clear

    For Each fila In tblBalance.ListRows

        fecha = fila.Range.Cells(1, _
            tblBalance.ListColumns("FECHA_BALANCE").Index).Value

        If IsDate(fecha) Then

            claveFecha = Format(CDate(fecha), "yyyymmdd")

            If Not fechas.Exists(claveFecha) Then

                fechas.Add claveFecha, True

                cboFechaBalance.AddItem _
                    Format(CDate(fecha), "dd/mm/yyyy")

            End If

        End If

    Next fila
End Sub
