Attribute VB_Name = "Módulo2"
Sub ImportarModuloDesdeEscritorio(FilePath As String)
    Dim VBComp As Object
    Dim VBP As Object
    Dim ModName As String
    Dim TargetBook As Workbook
    
    ' Nombre del módulo
    ModName = "NUMEROS A LETRAS"
    
    ' Referencia al libro de Excel activo
    Set TargetBook = ThisWorkbook
    
    ' Abrir el archivo .bas y añadirlo como un componente al proyecto actual
    Set VBComp = TargetBook.VBProject.VBComponents.Import(FilePath)
    
    ' Renombrar el componente VBA si ya existe
    On Error Resume Next
    VBComp.Name = ModName
    On Error GoTo 0
    
    MsgBox "El módulo ha sido importado correctamente.", vbInformation
End Sub
