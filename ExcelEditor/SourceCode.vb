Option Explicit
Sub ParseTable()
    'amount of all sheets
    Dim listCount As Integer
    listCount = Worksheets.Count
    'amount of students' sheets
    listCount = listCount - 1
    'first row of surnames index
    Dim nameCellIndex As Integer
    nameCellIndex = 14
    'first row of surnames letter
    Dim startNameCell As String
    startNameCell = "B"
    'Diploma number
    Dim diplomaCellValue As String
    'Practice grades start index(V4:V10)
    Dim practiceIndex As Integer
    practiceIndex = 4
    'Common grades 1st start index(V23:V38)
    Dim gradeIndex As Integer
    gradeIndex = 23
    'Common grades 2nd start index(V41:V76)
    Dim gradeIndex2 As Integer
    gradeIndex2 = 41
    'Student's name
    Dim nameCellValue As String
    'Main loop index
    Dim I As Integer
    'Practice grades loop index
    Dim pGradeCell As Range
    'Common grades 1st loop loop index
    Dim gradeCell As Range
    'Common grades 2nd loop index
    Dim gradeCell2 As Range
    
    'Remove all formulas
    Dim Worksheet As Worksheet, a, area As String
    For Each Worksheet In Worksheets
        a = Worksheet.UsedRange
        area = Worksheet.UsedRange.Address
        Worksheet.Cells.ClearContents
        Worksheet.Range(area) = a
    Next
    
    'Transfer data
    For I = 2 To listCount
        nameCellValue = Worksheets(1).Range("B" & CStr(nameCellIndex)).Value
        diplomaCellValue = Worksheets(1).Range("BP" & CStr(nameCellIndex)).Value
        Worksheets(I).Range("L5").Value = nameCellValue
        Worksheets(I).Range("T1").Value = diplomaCellValue
        'v4 - v10: Practice grades
        practiceIndex = 4
        For Each pGradeCell In Worksheets(1).Range("BC" & CStr(nameCellIndex) & ":BI" & CStr(nameCellIndex)).Cells
            Worksheets(I).Range("V" & CStr(practiceIndex)).Value = pGradeCell.Value
            practiceIndex = practiceIndex + 1
            If practiceIndex > 10 Then Exit For
        Next
        'State Exam
        Worksheets(I).Range("G19").Value = Worksheets(1).Range("BJ" & CStr(nameCellIndex)).Value
        'Grades C14:R14
        gradeIndex = 23
        For Each gradeCell In Worksheets(1).Range("C" & CStr(nameCellIndex) & ":R" & CStr(nameCellIndex)).Cells
            Worksheets(I).Range("V" & CStr(gradeIndex)).Value = gradeCell.Value
            gradeIndex = gradeIndex + 1
            Next
        'Grades S14:BB14
        gradeIndex2 = 41
        For Each gradeCell2 In Worksheets(1).Range("S" & CStr(nameCellIndex) & ":BB" & CStr(nameCellIndex)).Cells
            Worksheets(I).Range("V" & CStr(gradeIndex2)).Value = gradeCell2.Value
            gradeIndex2 = gradeIndex2 + 1
            Next
        
        'Registry number BQ14
        Worksheets(I).Range("H37").Value = Worksheets(1).Range("BQ" & CStr(nameCellIndex)).Value
        
        nameCellIndex = nameCellIndex + 1
        Next
End Sub
