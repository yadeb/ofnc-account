Sub FilterBankStatement()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, outRow As Long
    Dim validPurposes As Variant
    Dim i As Long

    ' === CONFIG ===
    Set wsSource = ThisWorkbook.Sheets("Statement") ' Source sheet
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("Filtered") ' Output sheet
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Sheets.Add(After:=wsSource)
        wsTarget.Name = "Filtered"
    Else
        wsTarget.Cells.Clear
    End If
    On Error GoTo 0

    validPurposes = Array("Groceries", "Rent", "Utilities", "Transport", "Savings") ' Allowed purposes
    outRow = 2 ' Start writing from row 2 (row 1 will have headers)

    ' === HEADERS ===
    wsTarget.Cells(1, 1).Value = "Purpose"
    wsTarget.Cells(1, 2).Value = "Amount"

    ' === FIND LAST ROW IN SOURCE ===
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' === LOOP AND COPY MATCHING ROWS ===
    For i = 2 To lastRow ' Assuming row 1 has headers
        Dim purpose As String
        purpose = Trim(wsSource.Cells(i, "B").Value) ' Purpose in column B
        If IsInArray(purpose, validPurposes) Then
            wsTarget.Cells(outRow, 1).Value = purpose
            wsTarget.Cells(outRow, 2).Value = wsSource.Cells(i, "C").Value ' Amount in column C
            outRow = outRow + 1
        End If
    Next i

    ' === CREATE TABLE ===
    Dim tblRange As Range
    Set tblRange = wsTarget.Range("A1").CurrentRegion
    Dim tbl As ListObject

    ' Remove existing table if present
    On Error Resume Next
    wsTarget.ListObjects("FilteredTable").Delete
    On Error GoTo 0

    Set tbl = wsTarget.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = "FilteredTable"
    tbl.TableStyle = "TableStyleMedium9"

    MsgBox "✅ Filtered data table created: 'FilteredTable' on 'Filtered' sheet.", vbInformation
End Sub

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim v
    For Each v In arr
        If StrComp(val, v, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next v
    IsInArray = False
End Function


⸻

' 📋 Assumptions:
' 	•	Your source sheet is named "Statement".
' 	•	Column B contains the Purpose, and Column C contains the Amount.
' 	•	You want the output on a sheet named "Filtered" (it creates this sheet if it doesn’t exist).
' 	•	Output is stored as an Excel Table named "FilteredTable".

' ⸻

' 🧠 How to Use It:
' 	1.	Press Alt + F11 to open the VBA editor.
' 	2.	Insert a new module: Insert > Module.
' 	3.	Paste the code.
' 	4.	Close the editor.
' 	5.	Press Alt + F8, run FilterBankStatement.

' ⸻

' Would you like it to also include totals per purpose, or make it refresh automatically with a button?