VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBalance 
   Caption         =   "Generar balance"
   ClientHeight    =   3520
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4480
   OleObjectBlob   =   "frmBalance.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboEmpresa_Change()
   CargarFechasBalance
   CargarFechasInspeccion
End Sub

Private Sub cboEspecie_Change()
   CargarFechasBalance
   CargarFechasInspeccion
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardarBalance_Click()
    Dim empresa As String
    Dim especie As String
    Dim fechaBalance As Date
    Dim fechaInspeccion As Date
    
    Dim tieneInspeccion As Boolean
    
    Dim m3Inspeccionado As Double
    Dim m3InventarioFisico As Double
    Dim entradasM3 As Double
    Dim salidasM3 As Double
    Dim balanceM3 As Double
    
    Dim tblInventarioBase As ListObject
    Dim filaInventarioBase As ListRow
    Dim tblInventarioFisico As ListObject
    Dim filaInventarioFisico As ListRow
    
    Dim tblEmpresas As ListObject
    Dim filaEmpresa As ListRow
    Dim idEmpresa As Long
    Dim tblDetalleMov As ListObject
    Dim filaDetalleMov As ListRow
    Dim fechaMovimiento As Date
    Dim movimientoEnRango As Boolean
    
    Dim tblBalance As ListObject
    Dim filaBalance As ListRow
    Dim balanceExiste As Boolean
    
    Dim tblEspecies As ListObject
    Dim filaEspecie As ListRow
    Dim idEspecie As Long
    
    If Not CampoRequerido(cboEmpresa.Value, "Empresa") Then Exit Sub
    If Not CampoRequerido(cboEspecie.Value, "Especie") Then Exit Sub
    If Not CampoRequerido(cboFechaBalance.Value, "Fecha balance") Then Exit Sub
    If Not FechaValida(cboFechaBalance.Value, "Fecha balance") Then Exit Sub
    If Not CampoRequerido(cboFechaInspeccion.Value, "Fecha inspección") Then Exit Sub
    
    If cboFechaInspeccion.Value <> "SIN INSPECCIÓN PREVIA" Then
        If Not FechaValida(cboFechaInspeccion.Value, "Fecha inspección") Then Exit Sub
        tieneInspeccion = True

    Else

        tieneInspeccion = False
    End If
    
    empresa = cboEmpresa.Value
    especie = cboEspecie.Value
    fechaBalance = CDate(cboFechaBalance.Value)
    
    Set tblInventarioBase = ThisWorkbook.Worksheets("INVENTARIO_BASE").ListObjects("tblInventarioBase")
    Set tblInventarioFisico = ThisWorkbook.Worksheets("INVENTARIO_FISICO").ListObjects("tblInventarioFisico")
    Set tblEmpresas = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEmpresas")
    Set tblDetalleMov = ThisWorkbook.Worksheets("DETALLE_MOV").ListObjects("tblDetalleMov")
    Set tblBalance = ThisWorkbook.Worksheets("BALANCE").ListObjects("tblBalance")
    Set tblEspecies = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEspecies")
    If tieneInspeccion Then
        fechaInspeccion = CDate(cboFechaInspeccion.Value)
        For Each filaInventarioBase In tblInventarioBase.ListRows
            If empresa = filaInventarioBase.Range.Columns(tblInventarioBase.ListColumns("EMPRESA").Index).Value And especie = filaInventarioBase.Range.Columns(tblInventarioBase.ListColumns("ESPECIE").Index).Value _
            And fechaInspeccion = filaInventarioBase.Range.Columns(tblInventarioBase.ListColumns("FECHA_INSPECCION").Index).Value Then
                m3Inspeccionado = filaInventarioBase.Range.Columns(tblInventarioBase.ListColumns("M3_INSPECCIONADO").Index).Value
                Exit For
            End If
        Next filaInventarioBase
    Else
        m3Inspeccionado = 0
    End If
    m3InventarioFisico = 0
    For Each filaInventarioFisico In tblInventarioFisico.ListRows
        If empresa = filaInventarioFisico.Range.Columns(tblInventarioFisico.ListColumns("EMPRESA").Index).Value And especie = filaInventarioFisico.Range.Columns(tblInventarioFisico.ListColumns("ESPECIE").Index).Value _
        And fechaBalance = filaInventarioFisico.Range.Columns(tblInventarioFisico.ListColumns("FECHA_INVENTARIO").Index).Value Then
            m3InventarioFisico = m3InventarioFisico + CDbl(filaInventarioFisico.Range.Columns(tblInventarioFisico.ListColumns("M3").Index).Value)
        End If
    Next filaInventarioFisico
    
    For Each filaEmpresa In tblEmpresas.ListRows
        If empresa = filaEmpresa.Range.Columns(tblEmpresas.ListColumns("EMPRESA").Index).Value Then
            idEmpresa = filaEmpresa.Range.Columns(tblEmpresas.ListColumns("ID_EMPRESA").Index).Value
            Exit For
        End If
    Next filaEmpresa
    If idEmpresa = 0 Then
        MsgBox "No se encontró una empresa válida.", vbExclamation
        Exit Sub
    End If
    
    For Each filaEspecie In tblEspecies.ListRows
        If especie = filaEspecie.Range.Columns(tblEspecies.ListColumns("ESPECIE").Index).Value Then
            idEspecie = filaEspecie.Range.Columns(tblEspecies.ListColumns("ID_ESPECIE").Index).Value
            Exit For
        End If
    Next filaEspecie
    
    If idEspecie = 0 Then
        MsgBox "No se encontró una especie válida.", vbExclamation
        Exit Sub
    End If
    
    entradasM3 = 0
    salidasM3 = 0
    'For oara recorrer tabla detalle mov
    For Each filaDetalleMov In tblDetalleMov.ListRows
        fechaMovimiento = CDate(filaDetalleMov.Range.Columns(tblDetalleMov.ListColumns("FECHA_MOVIMIENTO_AUX").Index).Value)
        If filaDetalleMov.Range.Columns(tblDetalleMov.ListColumns("ID_EMPRESA_AUX").Index).Value = idEmpresa _
        And filaDetalleMov.Range.Columns(tblDetalleMov.ListColumns("ESPECIE").Index).Value = especie _
        And filaDetalleMov.Range.Columns(tblDetalleMov.ListColumns("ESTADO_MOVIMIENTO_AUX").Index).Value <> "CANCELADA" Then
             movimientoEnRango = False
             If tieneInspeccion Then
                 If fechaMovimiento >= fechaInspeccion _
                 And fechaMovimiento <= fechaBalance Then
                    movimientoEnRango = True
                 End If
             Else
                 If fechaMovimiento <= fechaBalance Then
                    movimientoEnRango = True
                 End If
             End If
             If movimientoEnRango Then
                If filaDetalleMov.Range.Columns(tblDetalleMov.ListColumns("TIPO_MOVIMIENTO_AUX").Index).Value = "Entrada" Then
                    entradasM3 = entradasM3 + CDbl(filaDetalleMov.Range.Columns(tblDetalleMov.ListColumns("CANTIDAD_M3").Index).Value)
                Else
                    If filaDetalleMov.Range.Columns(tblDetalleMov.ListColumns("TIPO_MOVIMIENTO_AUX").Index).Value = "Salida" Then
                        salidasM3 = salidasM3 + CDbl(filaDetalleMov.Range.Columns(tblDetalleMov.ListColumns("CANTIDAD_M3").Index).Value)
                    End If
                
                End If
             End If
        End If
    Next filaDetalleMov
    
    balanceM3 = (m3Inspeccionado + entradasM3) - (m3InventarioFisico + salidasM3)
    
    balanceExiste = False
    
    For Each filaBalance In tblBalance.ListRows
        If filaBalance.Range.Columns( _
            tblBalance.ListColumns("ID_EMPRESA").Index _
        ).Value = idEmpresa _
        And filaBalance.Range.Columns( _
            tblBalance.ListColumns("ID_ESPECIE").Index _
        ).Value = idEspecie _
        And filaBalance.Range.Columns( _
            tblBalance.ListColumns("FECHA_BALANCE").Index _
        ).Value = fechaBalance Then
        
            balanceExiste = True
            Exit For
        
        End If
    Next filaBalance
    If balanceExiste Then
        MsgBox "El balance para esta empresa, especie y fecha ya fue generado.", vbExclamation
        Exit Sub
    End If
    
    Set filaBalance = tblBalance.ListRows.Add
    
        filaBalance.Range.Columns( _
        tblBalance.ListColumns("ID_EMPRESA").Index _
    ).Value = idEmpresa
    
    filaBalance.Range.Columns( _
        tblBalance.ListColumns("ID_ESPECIE").Index _
    ).Value = idEspecie
    
    filaBalance.Range.Columns( _
        tblBalance.ListColumns("FECHA_BALANCE").Index _
    ).Value = fechaBalance
    
    If tieneInspeccion Then
        filaBalance.Range.Columns( _
            tblBalance.ListColumns("FECHA_BASE").Index _
        ).Value = fechaInspeccion
    Else
        filaBalance.Range.Columns( _
            tblBalance.ListColumns("FECHA_BASE").Index _
        ).ClearContents
    End If
    
    filaBalance.Range.Columns( _
        tblBalance.ListColumns("M3_BASE").Index _
    ).Value = m3Inspeccionado
    
    filaBalance.Range.Columns( _
        tblBalance.ListColumns("ENTRADAS_M3").Index _
    ).Value = entradasM3
    
    filaBalance.Range.Columns( _
        tblBalance.ListColumns("SALIDAS_M3").Index _
    ).Value = salidasM3
    
    filaBalance.Range.Columns( _
        tblBalance.ListColumns("FECHA_INVENTARIO_FISICO").Index _
    ).Value = fechaBalance
    
    filaBalance.Range.Columns( _
        tblBalance.ListColumns("INVENTARIO_FISICO_M3").Index _
    ).Value = m3InventarioFisico
    
    filaBalance.Range.Columns( _
        tblBalance.ListColumns("BALANCE_M3").Index _
    ).Value = balanceM3
End Sub

Private Sub UserForm_Initialize()
    Dim tblEmpresas As ListObject
    Dim filaEmpresa As ListRow
    Dim tblEspecies As ListObject
    Dim filaEspecie As ListRow
    
    Set tblEmpresas = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEmpresas")
    Set tblEspecies = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEspecies")
    Set tblDetalleMov = ThisWorkbook.Worksheets("DETALLE_MOV").ListObjects("tblDetalleMov")
    
    For Each filaEmpresa In tblEmpresas.ListRows
        cboEmpresa.AddItem filaEmpresa.Range.Columns(tblEmpresas.ListColumns("EMPRESA").Index).Value
    Next filaEmpresa
    
    For Each filaEspecie In tblEspecies.ListRows
        cboEspecie.AddItem filaEspecie.Range.Columns(tblEspecies.ListColumns("ESPECIE").Index).Value
    Next filaEspecie
    
End Sub

Private Sub CargarFechasInspeccion()
    Dim tblInventarioBase As ListObject
    Dim filaInventarioBase As ListRow
    Dim fechasInspeccionAgregadas As Object
    Dim fechaInspeccion As Variant
    Dim hayInspeccion As Boolean
    
    hayInspeccion = False
    cboFechaInspeccion.Clear

    'cboFechaInspeccion.AddItem "SIN INSPECCIÓN PREVIA"
    
    Set fechasInspeccionAgregadas = CreateObject("Scripting.Dictionary")
    Set tblInventarioBase = ThisWorkbook.Worksheets("INVENTARIO_BASE").ListObjects("tblInventarioBase")
    
    For Each filaInventarioBase In tblInventarioBase.ListRows
        If filaInventarioBase.Range.Columns(tblInventarioBase.ListColumns("EMPRESA").Index).Value = cboEmpresa.Value And filaInventarioBase.Range.Columns(tblInventarioBase.ListColumns("ESPECIE").Index).Value = cboEspecie.Value Then
            fechaInspeccion = filaInventarioBase.Range.Columns(tblInventarioBase.ListColumns("FECHA_INSPECCION").Index).Value
            If Not fechasInspeccionAgregadas.Exists(fechaInspeccion) Then
                cboFechaInspeccion.AddItem fechaInspeccion
                fechasInspeccionAgregadas.Add fechaInspeccion, True
                hayInspeccion = True
            End If
        End If
    Next filaInventarioBase
    If Not hayInspeccion Then

        cboFechaInspeccion.AddItem "SIN INSPECCIÓN PREVIA"
    
    End If
    
End Sub
Private Sub CargarFechasBalance()
    Dim tblInventarioFisico As ListObject
    Dim filaInventario As ListRow
    Dim fechasAgregadas As Object
    Dim fechaInventario As Variant

    cboFechaBalance.Clear

    Set fechasAgregadas = CreateObject("Scripting.Dictionary")

    Set tblInventarioFisico = ThisWorkbook.Worksheets("INVENTARIO_FISICO") _
        .ListObjects("tblInventarioFisico")

    For Each filaInventario In tblInventarioFisico.ListRows

        If filaInventario.Range.Columns( _
            tblInventarioFisico.ListColumns("EMPRESA").Index _
        ).Value = cboEmpresa.Value And filaInventario.Range.Columns( _
        tblInventarioFisico.ListColumns("ESPECIE").Index _
        ).Value = cboEspecie.Value Then

            fechaInventario = filaInventario.Range.Columns( _
                tblInventarioFisico.ListColumns("FECHA_INVENTARIO").Index _
            ).Value

            If Not fechasAgregadas.Exists(fechaInventario) Then

                cboFechaBalance.AddItem fechaInventario
                fechasAgregadas.Add fechaInventario, True

            End If

        End If

    Next filaInventario
End Sub
