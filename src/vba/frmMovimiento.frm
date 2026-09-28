VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMovimiento 
   Caption         =   "Nuevo movimiento"
   ClientHeight    =   6040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11100
   OleObjectBlob   =   "frmMovimiento.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private indiceDetalleEditando As Long
Private Sub Label1_Click()

End Sub
    
Private Sub cboEmpresa_Change()
    ActualizarCamposForestales
End Sub

Private Sub cboTipoDetalle_Change()
    ActualizarCamposDetalle
End Sub

Private Sub cboTipoMovimiento_Change()
    ActualizarCamposForestales
End Sub

Private Sub cmdAgregarDetalle_Click()
    If Not CampoRequerido(cboTipoDetalle.Value, "Tipo de detalle") Then Exit Sub
    
    If cboTipoDetalle.Value = "MADERA" Then
        If Not CampoRequerido(cboEspecie.Value, "Especie") Then Exit Sub
        If Not CampoRequerido(txtCantidadM3.Value, "Cantidad m³") Then Exit Sub
        If Not NumeroPositivo(txtCantidadM3.Value, "Cantidad m³") Then Exit Sub
    ElseIf cboTipoDetalle.Value = "PRODUCTO_TERMINADO" Then
        If Not CampoRequerido(txtProductoTerminado.Value, "Producto terminado") Then Exit Sub
    End If
    
    If indiceDetalleEditando >= 0 Then
        'MsgBox "Índice en edición: " & indiceDetalleEditando
        lstDetalles.List(indiceDetalleEditando, 0) = cboTipoDetalle.Value
        If cboTipoDetalle.Value = "MADERA" Then
            lstDetalles.List(indiceDetalleEditando, 1) = cboEspecie.Value
            lstDetalles.List(indiceDetalleEditando, 2) = CDbl(txtCantidadM3.Value)
            lstDetalles.List(indiceDetalleEditando, 3) = ""
        ElseIf cboTipoDetalle.Value = "PRODUCTO_TERMINADO" Then
            lstDetalles.List(indiceDetalleEditando, 1) = ""
            lstDetalles.List(indiceDetalleEditando, 2) = ""
            lstDetalles.List(indiceDetalleEditando, 3) = txtProductoTerminado.Value
        End If
        
    Else
        If cboTipoDetalle.Value = "MADERA" Then
            lstDetalles.AddItem cboTipoDetalle.Value
            lstDetalles.List(lstDetalles.ListCount - 1, 1) = cboEspecie.Value
            lstDetalles.List(lstDetalles.ListCount - 1, 2) = CDbl(txtCantidadM3.Value)
            lstDetalles.List(lstDetalles.ListCount - 1, 3) = ""
        ElseIf cboTipoDetalle.Value = "PRODUCTO_TERMINADO" Then
             ' aquí pondremos la lógica para producto terminado
            lstDetalles.AddItem cboTipoDetalle.Value
            lstDetalles.List(lstDetalles.ListCount - 1, 1) = ""
            lstDetalles.List(lstDetalles.ListCount - 1, 2) = ""
            lstDetalles.List(lstDetalles.ListCount - 1, 3) = txtProductoTerminado.Value
        End If
    End If
    LimpiarDetalle
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCancelarEdicionDetalle_Click()
    LimpiarDetalle
    
End Sub

Private Sub cmdEditarDetalle_Click()
    If lstDetalles.ListIndex = -1 Then
        MsgBox "Seleccione un detalle para editar.", vbExclamation
    Exit Sub
    End If
    indiceDetalleEditando = lstDetalles.ListIndex
    cboTipoDetalle.Value = lstDetalles.List(indiceDetalleEditando, 0)
    If cboTipoDetalle.Value = "MADERA" Then
        cboEspecie.Value = lstDetalles.List(indiceDetalleEditando, 1)
        txtCantidadM3.Value = lstDetalles.List(indiceDetalleEditando, 2)
    ElseIf cboTipoDetalle.Value = "PRODUCTO_TERMINADO" Then
        txtProductoTerminado.Value = lstDetalles.List(indiceDetalleEditando, 3)
    End If
    cmdAgregarDetalle.Caption = "Guardar cambios"
End Sub

Private Sub cmdGuardar_Click()
    Dim tblMovimientos As ListObject
    Dim nuevaFila As ListRow
    Dim mensajeError As String
    Dim tblDetalleMov As ListObject
    Dim nuevaFilaDetalle As ListRow
    Dim idMovimiento As Long
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ManejoError
    
    If Not CampoRequerido(cboEmpresa.Value, "Empresa") Then Exit Sub
    If Not CampoRequerido(cboTipoMovimiento.Value, "Tipo de movimiento") Then Exit Sub
    If Not CampoRequerido(cboEstadoMovimiento.Value, "Estado del movimiento") Then Exit Sub
    If Not CampoRequerido(txtNoRemision.Value, "No. remisión") Then Exit Sub
    If Not FechaValida(txtFechaVencimiento.Value, "Fecha de vencimiento") Then Exit Sub
    'If Not CampoRequerido(txtFolioAutorizado.Value, "Folio autorizado") Then Exit Sub
    If cboEmpresa.Value = "ASERRADERO_DEMO" And cboTipoMovimiento.Value = "Entrada" Then
        If Not CampoRequerido(txtRemitente.Value, "Remitente") Then Exit Sub
        If Not CampoRequerido(txtCodigoIdentificacion.Value, "Código de identificación") Then Exit Sub
        If Not FechaValida(txtFecha.Value, "Fecha") Then Exit Sub
    End If
    
    If lstDetalles.ListCount = 0 Then

        MsgBox "Debe agregar al menos un detalle al movimiento.", vbExclamation
        Exit Sub
    
    End If
    
    Set tblMovimientos = ThisWorkbook.Worksheets("MOVIMIENTOS").ListObjects("tblMovimientos")
    Set tblDetalleMov = ThisWorkbook.Worksheets("DETALLE_MOV").ListObjects("tblDetalleMov")
    Set nuevaFila = tblMovimientos.ListRows.Add
    
    'se ocupo para una prueba controlada Err.Raise vbObjectError + 1000, , "Error de prueba para comprobar rollback"
    
    nuevaFila.Range.Columns(tblMovimientos.ListColumns("EMPRESA").Index).Value = cboEmpresa.Value
    nuevaFila.Range.Columns(tblMovimientos.ListColumns("TIPO_MOVIMIENTO").Index).Value = cboTipoMovimiento.Value
    If txtFecha.Enabled Then
        nuevaFila.Range.Columns( _
            tblMovimientos.ListColumns("FECHA").Index _
        ).Value = CDate(txtFecha.Value)
    
    Else
        nuevaFila.Range.Columns( _
            tblMovimientos.ListColumns("FECHA").Index _
        ).Value = ""
    End If
    'nuevaFila.Range.Columns(tblMovimientos.ListColumns("FECHA").Index).Value = IIf(txtFecha.Enabled, CDate(txtFecha.Value), "")
    nuevaFila.Range.Columns(tblMovimientos.ListColumns("NO_REMISION").Index).Value = txtNoRemision.Value
    nuevaFila.Range.Columns(tblMovimientos.ListColumns("FOLIO_AUTORIZADO").Index).Value = txtFolioAutorizado.Value
    nuevaFila.Range.Columns(tblMovimientos.ListColumns("REMITENTE").Index).Value = txtRemitente.Value
    'nuevaFila.Range.Columns(tblMovimientos.ListColumns("REMITENTE").Index).Value = txtRemitente.Value
    nuevaFila.Range.Columns(tblMovimientos.ListColumns("CODIGO_IDENTIFICACION").Index).Value = txtCodigoIdentificacion.Value
    nuevaFila.Range.Columns(tblMovimientos.ListColumns("FECHA_VENCIMIENTO").Index).Value = CDate(txtFechaVencimiento.Value)
    nuevaFila.Range.Columns(tblMovimientos.ListColumns("ESTADO_MOVIMIENTO").Index).Value = cboEstadoMovimiento.Value
    nuevaFila.Range.Columns(tblMovimientos.ListColumns("OBSERVACIONES").Index).Value = txtObservaciones.Value
    idMovimiento = CLng(nuevaFila.Range.Columns(tblMovimientos.ListColumns("ID_MOVIMIENTO").Index).Value)
    For i = 0 To lstDetalles.ListCount - 1
        ' Aquí guardaremos cada detalle
        Set nuevaFilaDetalle = tblDetalleMov.ListRows.Add
        nuevaFilaDetalle.Range.Columns(tblDetalleMov.ListColumns("ID_MOVIMIENTO").Index).Value = idMovimiento
        nuevaFilaDetalle.Range.Columns(tblDetalleMov.ListColumns("TIPO_DETALLE").Index).Value = lstDetalles.List(i, 0)
        nuevaFilaDetalle.Range.Columns(tblDetalleMov.ListColumns("ESPECIE").Index).Value = lstDetalles.List(i, 1)
        nuevaFilaDetalle.Range.Columns(tblDetalleMov.ListColumns("CANTIDAD_M3").Index).Value = lstDetalles.List(i, 2)
        nuevaFilaDetalle.Range.Columns(tblDetalleMov.ListColumns("PRODUCTO_TERMINADO").Index).Value = lstDetalles.List(i, 3)
    Next i
    
    MsgBox "Movimiento guardado correctamente.", vbInformation
    LimpiarFormulario
    Exit Sub
ManejoError:
    mensajeError = Err.Description

    If idMovimiento > 0 And Not tblDetalleMov Is Nothing Then
        For j = tblDetalleMov.ListRows.Count To 1 Step -1
            If tblDetalleMov.ListRows(j).Range.Columns(tblDetalleMov.ListColumns("ID_MOVIMIENTO").Index).Value = idMovimiento Then
                tblDetalleMov.ListRows(j).Delete
            End If
        Next j
    End If
    If Not nuevaFila Is Nothing Then
        nuevaFila.Delete
    End If

    MsgBox "Ocurrió un error al guardar el movimiento: " & mensajeError, vbCritical
    
End Sub

Private Sub UserForm_Initialize()
    Dim tblEmpresas As ListObject
    Dim filaEmpresa As ListRow
    Dim tblEspecies As ListObject
    Dim filaEspecie As ListRow
    
    Set tblEmpresas = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEmpresas")
    Set tblEspecies = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEspecies")
    
    indiceDetalleEditando = -1
    
    For Each filaEmpresa In tblEmpresas.ListRows
        cboEmpresa.AddItem filaEmpresa.Range.Columns(tblEmpresas.ListColumns("EMPRESA").Index).Value
    Next filaEmpresa
    
    For Each filaEspecie In tblEspecies.ListRows
        cboEspecie.AddItem filaEspecie.Range.Columns(tblEspecies.ListColumns("ESPECIE").Index).Value
    Next filaEspecie
    
    cboTipoMovimiento.AddItem "Entrada"
    cboTipoMovimiento.AddItem "Salida"
    
    cboEstadoMovimiento.AddItem "VIGENTE"
    cboEstadoMovimiento.AddItem "CANCELADA"
    
    cboTipoDetalle.AddItem "MADERA"
    cboTipoDetalle.AddItem "PRODUCTO_TERMINADO"
    ActualizarCamposDetalle

End Sub
Private Sub ActualizarCamposForestales()
    If cboEmpresa.Value = "ASERRADERO_DEMO" And cboTipoMovimiento.Value = "Entrada" Then
        txtRemitente.Enabled = True
        txtCodigoIdentificacion.Enabled = True
        txtFecha.Enabled = True
    Else
        txtRemitente.Enabled = False
        txtCodigoIdentificacion.Enabled = False
        txtFecha.Enabled = False
        
        txtRemitente.Value = ""
        txtCodigoIdentificacion.Value = ""
        txtFecha.Value = ""
    End If
End Sub
Private Sub LimpiarFormulario()
    cboEmpresa.Value = ""
    cboTipoMovimiento.Value = ""
    txtFecha.Value = ""
    txtNoRemision.Value = ""
    txtFolioAutorizado.Value = ""
    txtRemitente.Value = ""
    txtCodigoIdentificacion.Value = ""
    txtFechaVencimiento.Value = ""
    cboEstadoMovimiento.Value = ""
    txtObservaciones.Value = ""
    lstDetalles.Clear
    ActualizarCamposForestales
End Sub
Private Sub ActualizarCamposDetalle()
    If cboTipoDetalle.Value = "MADERA" Then
        cboEspecie.Enabled = True
        txtCantidadM3.Enabled = True
        txtProductoTerminado.Enabled = False
        txtProductoTerminado.Value = ""
    ElseIf cboTipoDetalle.Value = "PRODUCTO_TERMINADO" Then
        cboEspecie.Enabled = False
        txtCantidadM3.Enabled = False
        txtProductoTerminado.Enabled = True
        cboEspecie.Value = ""
        txtCantidadM3.Value = ""
    Else
        cboEspecie.Enabled = False
        txtCantidadM3.Enabled = False
        txtProductoTerminado.Enabled = False
    
        cboEspecie.Value = ""
        txtCantidadM3.Value = ""
        txtProductoTerminado.Value = ""
    End If
End Sub
Private Sub LimpiarDetalle()
    cboTipoDetalle.Value = ""
    cboEspecie.Value = ""
    txtCantidadM3.Value = ""
    txtProductoTerminado.Value = ""
    
    indiceDetalleEditando = -1
    cmdAgregarDetalle.Caption = "Agregar detalle"
    
    ActualizarCamposDetalle
End Sub
Private Sub cmdEliminarDetalle_Click()
    If lstDetalles.ListIndex = -1 Then
        MsgBox "Seleccione un detalle para eliminar.", vbExclamation
        Exit Sub
    End If
    lstDetalles.RemoveItem lstDetalles.ListIndex
End Sub
