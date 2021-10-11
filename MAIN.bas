Function versionController()
    
    'Description: Value returning function to see if any changes have been made to the SQL code.
    'Created by: Charles Myint <charles.myint@cevalogistics.com>
    'Create Date: 08/19/2021
    
    Dim fPath As String
    Dim wb As Workbook
    
    fPath = "\\10.8.101.9\Production_Control\BW3\3NV_ESD_CODE\Version Controller\VBA_MODULES\Prod\SQL Version Controller.xlsx"
    
    Set wb = Workbooks.Open(fPath)
    
    versionController = wb.Sheets("Sheet1").Range("A1").Value
    
    wb.Close

End Function
Function svia_data() As Variant

    'Description: Returns an array of carriers back to the procedure that called it.
    'Created by: Charles Myint <charles.myint@cevalogistics.com>
    'Create Date: 08/19/2021
    
    Dim fPath As String
    Dim wb As Workbook
    Dim sviaLastRow As Long
    Dim cut_times() As Variant
    
    fPath = "\\10.8.101.9\Production_Control\BW3\3NV_ESD_CODE\SVIA\3NV_SVIA.xlsx"
    
    Set wb = Workbooks.Open(fPath)
    
    sviaLastRow = wb.Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row
    
    ReDim Preserve cut_times(1 To sviaLastRow, 1 To 2)
    
    For r = 1 To sviaLastRow
        For c = 1 To 2
            If c = 2 Then
                cut_times(r, c) = Format(Cells(r, c).Value, "HH:MM")
            Else
                cut_times(r, c) = Cells(r, c).Value
            End If
        Next c
    Next r
    
    wb.Close SaveChanges:=False
    
    svia_data = cut_times
    
End Function

Sub ChangeLog()
    
    'Tool: 3NV WIP Tool
    'Description: Gathers 3NV order data and provides an organized data set based off line of business, order status, and expected ship date.
    'Created by: Charles Myint <charles.myint@cevalogistics.com>
    'Create Date: 08/19/2021
    
    '----------------------------------------------------------------------'
    ' 1. All changes notes must be added in this module below as a comment.
    '       a. Please be organized & descriptive with your notes.
    '----------------------------------------------------------------------'
    
    '-------------Change Notes-------------'
    ' 08/19/2021 - Charles Myint - v1.0.2021: Initialization of the tool
    ' 08/20/2021 - Charles Myint - v1.0.2021: Added SQL code as a range value.
    ' 08/20/2021 - Charles Myint - v1.0.2021: Added in a function to test the SQL code for updates.
    ' 08/20/2021 - Charles Myint - v1.0.2021: Utilized arrays to add comments to the main data table.
    ' 09/02/2021 - Charles Myint - v1.0.2021: Added in function for a carrier tabel with cut-times. 
    ''''''''''''''''''''''''''''''''''''''''
    
End Sub

Sub GetSQL_Data()

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Description: Creates SQL query to gather BNA order data.                          '
    ' Creator: Charles Myint <Charles.Myint@CEVALogistics.com                           '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
    Dim wb As Workbook
    Dim sqlCode As String
    Dim Lastrow As Long

    Set wb = ThisWorkbook
    sqlCode = versionController
    Lastrow = wb.Sheets("MAIN").Range("A" & Rows.Count).End(xlUp).Row
    
    With wb.Sheets("MAIN").Range("A2:ZZ" & Lastrow)
        .ClearContents
    End With

    With wb.Connections("BNA").ODBCConnection
        .CommandText = sqlCode
        .BackgroundQuery = False
        .Refresh
        .CommandText = "To modify the SQL query please reach out to Production Control."
    End With

    Call PULL

    Sheets("MAIN").Activate

    Call create_Array
    
    Sheets("PIVOT").Activate
    ActiveSheet.PivotTables("MAIN_Pivot").PivotCache.Refresh
    ActiveSheet.PivotTables("SDS_Pivot").PivotCache.Refresh
    ActiveSheet.PivotTables("BD_Pivot").PivotCache.Refresh


End Sub

Sub create_Array()
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Description: A less memory intense way to apply PC comments to the MAIN data set. '
    '                                                                                   '
    ' Creator: Love Coding and Play via YouTube.                                        '
    ' Video: VBA For Loop Data Matching using Array.                                    '
    ' Modified: Charles Myint <Charles.Myint@CEVALogistics.com                          '
    '                                                                                   '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Application.ScreenUpdating = False
    
    Dim carrier_data As Variant, bnaData() As Variant, cmmntData() As Variant
    Dim cmmnt As Variant, esd As Variant, hold As Variant
    Dim wb As Workbook
    Dim mainLastRow As Long, mainLastCol As Long, commentLastRow As Long
    Dim bnaNumber As String, cmmntNumber As String
    Dim today As Date
    
    Set wb = ThisWorkbook

    mainLastRow = wb.Sheets("MAIN").Range("A" & Rows.Count).End(xlUp).Row
    mainLastCol = wb.Sheets("MAIN").Cells(1, Columns.Count).End(xlToLeft).Column
    commentLastRow = wb.Sheets("COMMENTS").Range("A" & Rows.Count).End(xlUp).Row
    today = Format(Now(), "M/DD/YYYY")
    
    carrier_data = svia_data ' Calls function to get carrier data and returns an array of that data

    ReDim Preserve bnaData(1 To mainLastRow, 1 To mainLastCol)
    ReDim Preserve cmmntData(1 To commentLastRow, 1 To 2)
    
    '''''''''''''''''''''''''''''''''''''''''''
    ' Creates array for the BNA main data set.'
    ' r = rows & c = columns.                 '
    '''''''''''''''''''''''''''''''''''''''''''
    
    For r = 1 To mainLastRow
        For c = 1 To mainLastCol
            bnaData(r, c) = wb.Sheets("MAIN").Cells(r, c)
        Next c
    Next r
 
    '''''''''''''''''''''''''''''''''''''''''''
    ' Creates array for the comment data set. '
    ' r = rows & c = columns.                 '
    '''''''''''''''''''''''''''''''''''''''''''
    
    For r = 1 To commentLastRow
        For c = 1 To 2 ' Modify if adding more columns
            cmmntData(r, c) = wb.Sheets("COMMENTS").Cells(r, c)
        Next c
    Next r
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Adds comments to the main BNA array              '
    ' a = row in BNA Array & b = row in Comment Array  '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    For a = 2 To UBound(bnaData)
        bnaNumber = bnaData(a, 1)

        For b = LBound(cmmntData) To UBound(cmmntData)
            cmmntNumber = cmmntData(b, 1)

            If cmmntNumber = bnaNumber Then
                bnaData(a, 2) = cmmntData(b, 2)
                Exit For
            End If
        Next b
    Next a

    ''''''''''''''''''''''''''''''''''''''''''''
    ' Determines if orders are shippable WIP   '
    ''''''''''''''''''''''''''''''''''''''''''''

    For r = 2 To UBound(bnaData)
        cmmnt = bnaData(r, 2)
        hold = bnaData(r, 25)

        If hold = 1 Or hold = 3 Or hold = 4 Or hold = 7 Then
            bnaData(r, 3) = "Hold"
        ElseIf IsError(cmmnt) Then
            If CVErr(xlErrNA) = bnaData(r, 2) Then
                bnaData(r, 3) = "WIP"
            End If
        ElseIf Left(cmmnt, 2) = "" Then
            bnaData(r, 3) = "WIP"
        ElseIf Left(cmmnt, 2) = "PS" Then
            bnaData(r, 3) = "Part Shortage"
        Else: bnaData(r, 3) = "WIP"
        End If
    Next r

    ''''''''''''''''''''''''''''''''''''''''''''
    ' Determines WIP bucket (ship by)          '
    ''''''''''''''''''''''''''''''''''''''''''''

    For r = 2 To UBound(bnaData)
        esd = bnaData(r, 15)
        cmmnt = bnaData(r, 2)

        If esd = "" Or esd = today Or Left(cmmnt, 2) = "PD" Then
            bnaData(r, 4) = "1_DAY"
        ElseIf esd = today + 1 Then
            bnaData(r, 4) = "2_DAY"
        ElseIf esd = today + 2 Then
            bnaData(r, 4) = "3_DAY"
        ElseIf esd = today + 3 Then
            bnaData(r, 4) = "4_DAY"
        ElseIf esd >= today + 4 Then
            bnaData(r, 4) = "5_DAY"
        Else
            bnaData(r, 4) = "PAST_DUE"
        End If
    Next r

    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Adds in carrier cut-times to the orders.    ' 
    ' Note: Will add in logic to calc exceptions. '
    '''''''''''''''''''''''''''''''''''''''''''''''

    For mainRow = 2 To UBound(bnaData)
        carrier = bnaData(mainRow, 13)
        country = bnaData(mainRow, 7)
        gln = bnaData(mainRow, 10)
        esd = Format(bnaData(mainRow, 15), "MM/DD/YYYY")
        dlDte = Format(bnaData(mainRow, 17), "MM/DD/YYYY")
        dlTime = Format(bnaData(mainRow, 17), "HH:MM")
        curDte = Format(Now(), "MM/DD/YYYY")


    While bnaData(mainRow, 5) = ""
        For sviaRow = 2 To UBound(carrier_data)
            If bnaData(mainRow, 16) = "3" Or bnaData(mainRow, 16) = ":" Or bnaData(mainRow, 16) = "F" Then
                bnaData(mainRow, 5) = "0:00"
            ElseIf carrier = carrier_data(sviaRow, 1) Then
                bnaData(mainRow, 5) = carrier_data(sviaRow, 2)
            End If
        Next sviaRow

        If bnaData(mainRow, 5) = "" Then
            bnaData(mainRow, 5) = "Needs Cut-time"
        End If
    Wend

    Next mainRow

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Places comments and if shippable wip to the main BNA main data sheet.                           '
    ' Note: Do not resize the entire range - this causes an issue with deleting the wb connection     '
    '       Add columns at the beginning of the table and redimension everything.                     '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Range("A:F").Resize(UBound(bnaData), 6).Value = bnaData

    Application.ScreenUpdating = True

End Sub
