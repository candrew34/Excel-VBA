Sub CreateTableOfContents()

    ' V3: Uses ActiveWorkbook to ensure the TOC is created in the file
    ' you have open, not the file where the macro is stored.

    ' Declare variables
    Dim wb As Workbook
    Dim tocSheet As Worksheet
    Dim ws As Worksheet
    Dim i As Long

    ' Set our target to be the currently open and active workbook
    Set wb = ActiveWorkbook

    ' Optimize macro speed
    Application.ScreenUpdating = False

    ' --- PREPARATION ---
    ' 1. Delete the old TOC sheet from the active workbook.
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("Table of Contents").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' 2. Add a new TOC sheet to the active workbook.
    Set tocSheet = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    tocSheet.Name = "Table of Contents"

    ' 3. Create a title for the TOC.
    With tocSheet.Range("A1")
        .Value = "Table of Contents"
        .Font.Bold = True
        .Font.Size = 16
        .Font.Underline = xlUnderlineStyleSingle
    End With

    ' --- MAIN LOOP ---
    ' Initialize the row counter.
    i = 3

    ' 4. Loop through every worksheet in the active workbook.
    For Each ws In wb.Worksheets
        If ws.Name <> tocSheet.Name Then
            
            ' Check if the cell is effectively blank (trims spaces).
            If Trim(ws.Range("A1").Value) = "" Then
                tocSheet.Cells(i, 1).Value = ws.Name
            Else
                tocSheet.Cells(i, 1).Value = ws.Range("A1").Value
            End If
            
            ' Create a hyperlink in that cell.
            tocSheet.Hyperlinks.Add Anchor:=tocSheet.Cells(i, 1), _
                                    Address:="", _
                                    SubAddress:="'" & ws.Name & "'!A1"
            
            ' Move to the next row.
            i = i + 1
        End If
    Next ws

    ' --- CLEANUP ---
    ' 5. Autofit the column width.
    tocSheet.Columns("A").AutoFit
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True

    ' Let the user know it's done.
    MsgBox "Table of Contents has been created successfully!", vbInformation

End Sub
