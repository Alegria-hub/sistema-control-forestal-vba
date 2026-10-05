VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCancelarMovimiento 
   Caption         =   "Cancelar movimiento"
   ClientHeight    =   3620
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5060
   OleObjectBlob   =   "frmCancelarMovimiento.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCancelarMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    Dim tblMovimientos As ListObject
    Dim filaMovimiento As ListRow
    Dim encontrado As Boolean
    
    Set tblMovimientos = ThisWorkbook.Worksheets("MOVIMIENTOS").ListObjects("tblMovimientos")
    
    encontrado = False
    
    For Each filaMovimiento In tblMovimientos.ListRows
        If Trim(CStr(txtNoRemision.Value)) = Trim(CStr( _
            filaMovimiento.Range.Columns( _
                tblMovimientos.ListColumns("NO_REMISION").Index _
            ).Value)) Then
                                txtEmpresa.Value = filaMovimiento.Range.Columns( _
                    tblMovimientos.ListColumns("EMPRESA").Index _
                ).Value
                
                txtTipoMovimiento.Value = filaMovimiento.Range.Columns( _
                    tblMovimientos.ListColumns("TIPO_MOVIMIENTO").Index _
                ).Value
                
                txtEstado.Value = filaMovimiento.Range.Columns( _
                    tblMovimientos.ListColumns("ESTADO_MOVIMIENTO").Index _
                ).Value
            MsgBox "Remisión encontrada.", vbInformation
            encontrado = True
            Exit For
        End If
    Next filaMovimiento
    If Not encontrado Then

        MsgBox "No se encontró la remisión.", vbExclamation
    
        txtEmpresa.Value = ""
        txtTipoMovimiento.Value = ""
        txtEstado.Value = ""
    
    End If
End Sub

Private Sub cmdCancelarMovimiento_Click()
    Dim tblMovimientos As ListObject
    Dim filaMovimiento As ListRow
    Dim respuesta As VbMsgBoxResult

    If Trim(txtNoRemision.Value) = "" Then
        MsgBox "Primero busca una remisión.", vbExclamation
        Exit Sub
    End If

    If Trim(txtEstado.Value) = "" Then
        MsgBox "Primero busca una remisión válida.", vbExclamation
        Exit Sub
    End If

    If UCase(Trim(txtEstado.Value)) = "CANCELADA" Then
        MsgBox "Esta remisión ya está cancelada.", vbExclamation
        Exit Sub
    End If

    respuesta = MsgBox( _
        "¿Seguro que deseas cancelar la remisión " & txtNoRemision.Value & "?", _
        vbQuestion + vbYesNo, _
        "Confirmar cancelación" _
    )

    If respuesta = vbNo Then Exit Sub

    Set tblMovimientos = ThisWorkbook.Worksheets("MOVIMIENTOS").ListObjects("tblMovimientos")

    For Each filaMovimiento In tblMovimientos.ListRows

        If Trim(CStr(txtNoRemision.Value)) = Trim(CStr( _
            filaMovimiento.Range.Columns( _
                tblMovimientos.ListColumns("NO_REMISION").Index _
            ).Value)) Then

            filaMovimiento.Range.Columns( _
                tblMovimientos.ListColumns("ESTADO_MOVIMIENTO").Index _
            ).Value = "CANCELADA"

            txtEstado.Value = "CANCELADA"

            MsgBox "Remisión cancelada correctamente.", vbInformation
            Exit For

        End If

    Next filaMovimiento
End Sub

Private Sub cmdCerrar_Click()
 Unload Me
End Sub

