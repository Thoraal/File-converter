Sub PedirPIN()
    Dim pinIngresado As String
    Dim pinCorrecto As String
    
    ' Define el PIN correcto
    pinCorrecto = "1234"  ' Puedes cambiar esto por el PIN que desees
    
    ' Solicita al usuario que ingrese el PIN
    pinIngresado = InputBox("Ingrese el PIN:", "Verificación de PIN")
    
    ' Compara el PIN ingresado con el PIN correcto
    If pinIngresado = pinCorrecto Then
        MsgBox "PIN correcto. Acceso permitido.", vbInformation, "Acceso permitido"
        ' Aquí puedes llamar a la función o procedimiento que desees ejecutar después de verificar el PIN
    Else
        MsgBox "PIN incorrecto. Acceso denegado.", vbExclamation, "Acceso denegado"
        ' Puedes agregar aquí acciones adicionales si el PIN es incorrecto
    End If
End Sub

