Attribute VB_Name = "modValidaciones"

Public Function CampoRequerido(ByVal valor As String, ByVal nombreCampo As String) As Boolean
    If Trim(valor) = "" Then
        MsgBox "El campo " & nombreCampo & " es obligatorio.", vbExclamation
        CampoRequerido = False
        Exit Function
    End If
    CampoRequerido = True
End Function
Public Function FechaValida(ByVal valor As String, ByVal nombreCampo As String) As Boolean
    If Trim(valor) = "" Then
        MsgBox "El campo " & nombreCampo & " es obligatorio.", vbExclamation
        FechaValida = False
        Exit Function
    End If
    If Not IsDate(valor) Then
        MsgBox "El campo " & nombreCampo & " no contiene una fecha válida.", vbExclamation
        FechaValida = False
        Exit Function
    End If
    FechaValida = True
End Function
Public Function NumeroPositivo(ByVal valor As String, ByVal nombreCampo As String) As Boolean
    Dim numero As Double

    On Error GoTo NumeroInvalido

    numero = ValorDecimal(valor)

    If numero <= 0 Then
        MsgBox "El campo " & nombreCampo & " debe ser mayor que cero.", vbExclamation
        NumeroPositivo = False
        Exit Function
    End If

    NumeroPositivo = True
    Exit Function

NumeroInvalido:

    MsgBox "El campo " & nombreCampo & " debe contener un número válido.", vbExclamation
    NumeroPositivo = False

End Function
Public Function ValorDecimal(ByVal valor As Variant) As Double

    Dim texto As String
    Dim regex As Object

    texto = Trim(CStr(valor))
    texto = Replace(texto, " ", "")

    ' Aceptar punto o coma como decimal
    texto = Replace(texto, ",", ".")

    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "^[+-]?[0-9]+(\.[0-9]+)?$"

    If Not regex.Test(texto) Then
        Err.Raise 5
    End If

    ' Val siempre interpreta el punto como decimal
    ValorDecimal = Val(texto)


End Function
Public Function MedidaValida(ByVal valor As String) As Boolean
    Dim regex As Object
    Dim textoNormalizado As String
    Dim partes() As String
    Dim i As Long
    Dim parteActual As String
    Dim denominador As Long
    
    Set regex = CreateObject("VBScript.RegExp")
    textoNormalizado = UCase(Trim(valor))
    regex.Pattern = "^([0-9]+|[0-9]+/[0-9]+|[0-9]+ [0-9]+/[0-9]+)\s*X\s*([0-9]+|[0-9]+/[0-9]+|[0-9]+ [0-9]+/[0-9]+)\s*X\s*([0-9]+|[0-9]+/[0-9]+|[0-9]+ [0-9]+/[0-9]+)""?$"
    
    If Not regex.Test(textoNormalizado) Then

        MsgBox "La medida no tiene un formato válido.", vbExclamation
        MedidaValida = False
        Exit Function
    
    End If
    
    partes = Split(textoNormalizado, "X")
    
    For i = LBound(partes) To UBound(partes)
        parteActual = Trim(Replace(partes(i), """", ""))
        If InStr(parteActual, "/") > 0 Then
            denominador = CLng(Mid(parteActual, InStr(parteActual, "/") + 1))
            If denominador = 0 Then
                MsgBox "La medida contiene una fracción con denominador cero.", vbExclamation
                MedidaValida = False
                Exit Function
            End If
        End If
    Next i
    
    MedidaValida = True
End Function

Public Function NumeroNoNegativo(ByVal valor As Variant, ByVal nombreCampo As String) As Boolean

    If Not IsNumeric(valor) Then
        MsgBox nombreCampo & " debe ser un número.", vbExclamation
        NumeroNoNegativo = False
        Exit Function
    End If

    If CDbl(valor) < 0 Then
        MsgBox nombreCampo & " no puede ser negativo.", vbExclamation
        NumeroNoNegativo = False
        Exit Function
    End If

    NumeroNoNegativo = True

End Function
