Private ww()
    Dim pin As String
    Dim inputPin As String

    ' Define el PIN deseado
    pin = "1234"  ' Puedes cambiar esto por tu PIN deseado

    ' Pide al usuario que ingrese el PIN
    inputPin = InputBox("Ingrese el PIN:", "Autenticación")

    ' Comprueba si el PIN ingresado es correcto
    If inputPin = pin Then
        MsgBox "PIN correcto. Bienvenido.", vbInformation
    Else
        MsgBox "PIN incorrecto. Esta acción será registrada.", vbExclamation
        ThisWorkbook.Close savechanges:=False
    End If
End Sub
