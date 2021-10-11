Sub modImport()

    Dim wkbTarget As Excel.Workbook
    Dim cmpComponents As VBIDE.VBComponents
    Dim a As Integer
    
    Set wkbTarget = ThisWorkbook
    Set cmpComponents = wkbTarget.VBProject.VBComponents
    a = 1
    
    
    For a = 1 To cmpComponents.Count
        If cmpComponents(a).Name = "MAIN" Then
            GoTo run_code
        Else
            cmpComponents.Import("\\10.8.101.9\Production_Control\BW3\3NV_ESD_CODE\MAIN.bas").Name = "MAIN"
    Next

run_code:
    Call GetSQL_Data
    
End Sub

