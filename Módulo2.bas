Attribute VB_Name = "M�dulo2"
Sub ImportarModuloDesdeEscritorio(FilePath As String)
    Dim VBComp As Object
    Dim VBP As Object
    Dim ModName As String
    Dim TargetBook As Workbook
    
    ' Nombre del m�dulo
    ModName = "NUMEROS A LETRAS"
    
    ' Referencia al libro de Excel activo
    Set TargetBook = ThisWorkbook
    
    ' Abrir el archivo .bas y a�adirlo como un componente al proyecto actual
    Set VBComp = TargetBook.VBProject.VBComponents.Import(FilePath)
    
    ' Renombrar el componente VBA si ya existe
    On Error Resume Next
    VBComp.Name = ModName
    On Error GoTo 0
    
    MsgBox "El m�dulo ha sido importado correctamente.", vbInformation
End Sub
