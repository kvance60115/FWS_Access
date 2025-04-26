Sub importFWS()
    Dim db As DAO.Database
    Dim xlApp As New Excel.Application
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim arrRange As Range
    Dim lastrow As Long
    Dim firstrow As Long
    Dim firstcol As Long
    Dim lastcol As Long

    On Error GoTo importError ' Error handler for unexpected errors

    Set db = CurrentDb

    ' Prompt user to select an Excel file
    With Application.FileDialog(3)
        .Title = "Select Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No file selected. Exiting."
            GoTo cleanup
        End If
        filePath = .SelectedItems(1)
    End With

    ' Start Excel application
    xlApp.Visible = True

    ' Open the selected workbook
    Set wb = xlApp.Workbooks.Open(filePath, ReadOnly:=True)

    ' Reference the specific worksheet, isolate the error
    On Error Resume Next
    Set ws = wb.Worksheets("HRSDetail")
    On Error GoTo importError

    If ws Is Nothing Then
        MsgBox "Worksheet 'HRSDetail' not found."
        GoTo cleanup
    End If

    ' Wait to ensure workbook is ready (optional)
    cleanUpWorksheet ws
    DoEvents
    xlApp.Wait Now + TimeValue("0:00:01")

    ' Define range from C16:Z(last row)
    firstrow = 16
    firstcol = 3
    lastcol = 26
    lastrow = ws.Cells(ws.Rows.Count, firstcol).End(xlUp).Row

    Set arrRange = ws.Range(ws.Cells(firstrow, firstcol), ws.Cells(lastrow, lastcol))
    Debug.Print "Range is: " & arrRange.Address
   rangeToArr arrRange, xlApp
    'arrayToTable rangeToArr(arrRange, xlApp)
    MsgBox ("done")
cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Set db = Nothing
    Exit Sub

importError:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume cleanup
End Sub



Sub cleanUpWorksheet(ws As Object)

    ' Remove subtotals
    On Error Resume Next
    ws.Cells.RemoveSubtotal
    On Error GoTo 0

    ' Ungroup all (rows and columns)
    On Error Resume Next
    ws.Rows.Ungroup
    ws.Columns.Ungroup
    On Error GoTo 0

    ' Show all rows and columns
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False

    
End Sub

Function rangeToArr(ByRef r As Range, a As Excel.Application) As Variant
    Dim rowcount As Long, colcount As Long
    rowcount = r.Rows.Count
    colcount = 5
    
    Dim myArr() As Variant
    ReDim myArr(1 To rowcount, 1 To colcount)
    
    Dim tmpTotal As Double
    Dim currID As String, nextID As String
    Dim currName As String, currAcct As String, currProj As String
    Dim currAmt As Double
    
    Dim i As Long, j As Long
    j = 1
    
    For i = 1 To rowcount
        currID = Trim(CStr(r.Cells(i, 24).Value))
        If currID = "" Then GoTo SkipRow
        
        currName = r.Cells(i, 3)
        currAcct = r.Cells(i, 1)
        currProj = r.Cells(i, 19)
        currAmt = CDbl(r.Cells(i, 12).Value)
        tmpTotal = tmpTotal + currAmt
        
        ' Look ahead to next *non-blank* row for ID
        nextID = ""
        Dim lookAhead As Long
        For lookAhead = i + 1 To rowcount
            nextID = Trim(CStr(r.Cells(lookAhead, 24).Value))
            If nextID <> "" Then Exit For
        Next lookAhead
        
        If currID <> nextID Then
            myArr(j, 1) = currID
            myArr(j, 2) = currName
            myArr(j, 3) = tmpTotal
            myArr(j, 4) = currAcct
            myArr(j, 5) = currProj
            j = j + 1
            tmpTotal = 0
        End If

SkipRow:
    Next i

    ' Trim array
    Dim retArray() As Variant
    ReDim retArray(1 To j - 1, 1 To 5)
    
    Dim k As Long, m As Long
    For k = 1 To j - 1
        For m = 1 To 5
            retArray(k, m) = myArr(k, m)
        Next m
    Next k

    rangeToArr = retArray
End Function


Sub arrayToTable(ByRef r As Variant)

    Dim db As DAO.Database
    Set db = CurrentDb
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("Table1", dbOpenDynaset)
    
    Dim i As Long
    For i = 1 To UBound(r, 1)
        rs.AddNew
            rs.Fields(0).Value = r(i, 1)
            rs.Fields(1).Value = r(i, 2)
            rs.Fields(2).Value = r(i, 3)
            rs.Fields(3).Value = r(i, 4)
            rs.Fields(4).Value = r(i, 5)
        rs.Update
    Next i
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

