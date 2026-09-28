VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInventarioBase 
   Caption         =   "Captura de inventario base"
   ClientHeight    =   4380
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6700
   OleObjectBlob   =   "frmInventarioBase.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmInventarioBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Dim tblInventarioBase As ListObject
    Dim nuevaFila As ListRow
    Dim mensajeError As String

    On Error GoTo ManejoError
    
    If Not CampoRequerido(cboEmpresa.Value, "Empresa") Then Exit Sub
    If Not FechaValida(txtFechaInspeccion.Value, "Fecha de inspección") Then Exit Sub
    If Not CampoRequerido(cboEspecie.Value, "Especie") Then Exit Sub
    If Not CampoRequerido(txtM3Inspeccionado.Value, "M3 inspeccionado") Then Exit Sub
    If Not NumeroNoNegativo(txtM3Inspeccionado.Value, "M3 inspeccionado") Then Exit Sub
    
    Set tblInventarioBase = ThisWorkbook.Worksheets("INVENTARIO_BASE").ListObjects("tblInventarioBase")
    Set nuevaFila = tblInventarioBase.ListRows.Add
    
    nuevaFila.Range.Columns(tblInventarioBase.ListColumns("EMPRESA").Index).Value = cboEmpresa.Value
    nuevaFila.Range.Columns(tblInventarioBase.ListColumns("FECHA_INSPECCION").Index).Value = CDate(txtFechaInspeccion.Value)
    nuevaFila.Range.Columns(tblInventarioBase.ListColumns("ESPECIE").Index).Value = cboEspecie.Value
    nuevaFila.Range.Columns(tblInventarioBase.ListColumns("M3_INSPECCIONADO").Index).Value = CDbl(txtM3Inspeccionado.Value)
    nuevaFila.Range.Columns(tblInventarioBase.ListColumns("REFERENCIA").Index).Value = Trim(txtReferencia.Value)
    nuevaFila.Range.Columns(tblInventarioBase.ListColumns("OBSERVACIONES").Index).Value = Trim(txtObservaciones.Value)
    
    MsgBox "Inventario base guardado correctamente.", vbInformation
    LimpiarFormulario
    Exit Sub
ManejoError:

    mensajeError = Err.Description

    If Not nuevaFila Is Nothing Then
        nuevaFila.Delete
    End If

    MsgBox "Ocurrió un error al guardar el inventario base:" & vbCrLf & _
           mensajeError, vbCritical
End Sub
Private Sub LimpiarFormulario()

    cboEmpresa.Value = ""
    txtFechaInspeccion.Value = ""
    cboEspecie.Value = ""
    txtM3Inspeccionado.Value = ""
    txtReferencia.Value = ""
    txtObservaciones.Value = ""

End Sub

Private Sub UserForm_Initialize()
    Dim tblEmpresas As ListObject
    Dim filaEmpresa As ListRow
    Dim tblEspecies As ListObject
    Dim filaEspecie As ListRow
    
    Set tblEmpresas = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEmpresas")
    Set tblEspecies = ThisWorkbook.Worksheets("CATALOGOS").ListObjects("tblEspecies")
    
    For Each filaEmpresa In tblEmpresas.ListRows
        cboEmpresa.AddItem filaEmpresa.Range.Columns(tblEmpresas.ListColumns("EMPRESA").Index).Value
    Next filaEmpresa
    
    For Each filaEspecie In tblEspecies.ListRows
        cboEspecie.AddItem filaEspecie.Range.Columns(tblEspecies.ListColumns("ESPECIE").Index).Value
    Next filaEspecie
End Sub
