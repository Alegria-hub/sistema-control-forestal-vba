VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInventarioFisico 
   Caption         =   "Captura de inventario físico"
   ClientHeight    =   4210
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7280
   OleObjectBlob   =   "frmInventarioFisico.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmInventarioFisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Dim tblInventarioFisico As ListObject
    Dim nuevaFila As ListRow
    Dim mensajeError As String
    
    On Error GoTo ManejoError
    
    If Not CampoRequerido(cboEmpresa.Value, "Empresa") Then Exit Sub
    If Not CampoRequerido(cboEspecie.Value, "Especie") Then Exit Sub
    If Not FechaValida(txtFechaInventario.Value, "Fecha de inventario") Then Exit Sub
    If Not CampoRequerido(cboEncargado.Value, "Encargado") Then Exit Sub
    If Not CampoRequerido(cboTipoProducto.Value, "Tipo de producto") Then Exit Sub
    If Not CampoRequerido(txtCantidad.Value, "Cantidad") Then Exit Sub
    If Not NumeroPositivo(txtCantidad.Value, "Cantidad") Then Exit Sub
    If Not CampoRequerido(txtMedida.Value, "Medida") Then Exit Sub
    If Not MedidaValida(txtMedida.Value) Then Exit Sub
        
    Set tblInventarioFisico = ThisWorkbook.Worksheets("INVENTARIO_FISICO").ListObjects("tblInventarioFisico")
    
    Set nuevaFila = tblInventarioFisico.ListRows.Add
    
    nuevaFila.Range.Columns(tblInventarioFisico.ListColumns("EMPRESA").Index).Value = cboEmpresa.Value
    nuevaFila.Range.Columns(tblInventarioFisico.ListColumns("ESPECIE").Index).Value = cboEspecie.Value
    nuevaFila.Range.Columns(tblInventarioFisico.ListColumns("FECHA_INVENTARIO").Index).Value = CDate(txtFechaInventario.Value)
    nuevaFila.Range.Columns(tblInventarioFisico.ListColumns("ENCARGADO").Index).Value = cboEncargado.Value
    nuevaFila.Range.Columns(tblInventarioFisico.ListColumns("TIPO_PRODUCTO").Index).Value = cboTipoProducto.Value
    nuevaFila.Range.Columns(tblInventarioFisico.ListColumns("CANTIDAD").Index).Value = CDbl(txtCantidad.Value)
    nuevaFila.Range.Columns(tblInventarioFisico.ListColumns("MEDIDA").Index).Value = UCase(Trim(txtMedida.Value))
    
    MsgBox "Inventario físico guardado correctamente.", vbInformation
    LimpiarFormulario
    
    Exit Sub

ManejoError:
    mensajeError = Err.Description

    If Not nuevaFila Is Nothing Then
        nuevaFila.Delete
    End If
    
    MsgBox "Ocurrió un error al guardar el inventario físico: " & mensajeError, vbCritical
    
    
End Sub

Private Sub UserForm_Initialize()
    Dim tblEmpresas As ListObject
    Dim filaEmpresa As ListRow
    Dim tblEspecies As ListObject
    Dim filaEspecie As ListRow
    Dim tblEncargados As ListObject
    Dim filaEncargado As ListRow
    
    Set tblEmpresas = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEmpresas")
    Set tblEspecies = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEspecies")
    Set tblEncargados = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEncargados")
    
    For Each filaEmpresa In tblEmpresas.ListRows
        cboEmpresa.AddItem filaEmpresa.Range.Columns(tblEmpresas.ListColumns("EMPRESA").Index).Value
    Next filaEmpresa
    
    For Each filaEspecie In tblEspecies.ListRows
        cboEspecie.AddItem filaEspecie.Range.Columns(tblEspecies.ListColumns("ESPECIE").Index).Value
    Next filaEspecie
    
    For Each filaEncargado In tblEncargados.ListRows
        cboEncargado.AddItem filaEncargado.Range.Columns(tblEncargados.ListColumns("ENCARGADO").Index).Value
    Next filaEncargado
    
    cboTipoProducto.AddItem "VIGA"
    cboTipoProducto.AddItem "POLIN"
    cboTipoProducto.AddItem "TABLA"
    cboTipoProducto.AddItem "TABLETA"
End Sub
Private Sub LimpiarFormulario()
    cboEmpresa.Value = ""
    cboEspecie.Value = ""
    txtFechaInventario.Value = ""
    cboEncargado.Value = ""
    cboTipoProducto.Value = ""
    txtCantidad.Value = ""
    txtMedida.Value = ""
End Sub
