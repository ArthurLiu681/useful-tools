Attribute VB_Name = "Module12"
Sub write_Word()
Attribute write_Word.VB_ProcData.VB_Invoke_Func = "W\n14"
Dim code As Integer
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim t As Table
Dim benchline As String
Dim mat As Long
Dim pm As Long
Dim cct As Long
Dim pmat As Long
Dim ppm As Long
Dim pcct As Long
Dim my_excel As Workbook
Dim wordapp As Object
For i = 101 To 120
    Workbooks("PERSONAL.XLSB").Activate
    code = Cells(i, 1).Value
    AppActivate "Sun Hung Kai Properties Ltd - Others Segment [S1083_DQC20-0034-01-00_A01] - Engagement Management System"
    Application.Wait Now + TimeValue("0:00:01")
    SendKeys "{LEFT}"
    Application.Wait Now + TimeValue("0:00:01")
    SendKeys "{DOWN}"
    Application.Wait Now + TimeValue("0:00:01")
    SendKeys "{DOWN}"
    Application.Wait Now + TimeValue("0:00:01")
    SendKeys "+{DOWN}"
    Application.Wait Now + TimeValue("0:00:01")
    SendKeys "{ENTER}"
    Application.Wait Now + TimeValue("0:00:01")
    For Each Workbook In Workbooks
        If Workbook.Name Like "*Determine*" Then
            Set my_excel = Workbook
            my_excel.Activate
            mat = Range("B27").Value
            pm = Range("B29").Value
            cct = Range("B35").Value
            benchline = Range("B17").Value
            Workbooks("2nd round 2021 Combined TB.xlsx").Activate
            c = Cells.Find(What:=code, After:=Cells(1, 1), LookIn:=xlFormulas, LookAt _
                    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                    False, SearchFormat:=False).Column
            If benchline = "Revenue" Then
                benchmark = Abs(Application.WorksheetFunction.Sum(Range(Cells(761, c).Address & ":" & Cells(768, c).Address)))
            Else
                If benchline = "Total assets" Then
                    benchmark = Abs(Application.WorksheetFunction.Sum(Range(Cells(2, c).Address & ":" & Cells(539, c).Address)))
                Else
                    benchmark = Abs(Application.WorksheetFunction.Sum(Range(Cells(2, c).Address & ":" & Cells(623, c).Address)) + Application.WorksheetFunction.Sum(Range(Cells(642, c).Address & ":" & Cells(661, c).Address)))
                End If
            End If
            my_excel.Activate
            Range("B19") = benchmark
            pmat = Range("B27").Value
            ppm = Range("B29").Value
            pcct = Range("B35").Value
            Exit For
        End If
    Next Workbook
    Set wrdapp = GetObject(, "Word.Application")
ErrResume:
    On Error GoTo ErrPaste
        wrdapp.Documents("13900 Comprehensive audit planning memorandum without EMS links_.docx").Activate
    On Error GoTo 0
ErrPaste:
    If Err.Number = 462 Then
        Set wrdapp = CreateObject("Word.Application")
        Resume ErrResume
    End If
    Set t = wrdapp.ActiveDocument.Tables(10)
    t.Cell(2, 1).Range.Text = mat
    t.Cell(2, 2).Range.Text = pm
    t.Cell(2, 3).Range.Text = cct
    t.Cell(2, 4).Range.Text = pmat
    t.Cell(2, 5).Range.Text = ppm
    t.Cell(2, 6).Range.Text = pcct
    For j = 1 To 6
        t.Cell(2, j).Range.HighlightColorIndex = wdNoHighlight
    Next j
    wrdapp.ActiveDocument.Close savechanges:=wdSaveChanges
    wrdapp.Quit
    my_excel.Close savechanges:=False
    Set my_excel = Nothing
    If Not (wrdapp Is Nothing) Then
        Set wrdapp = Nothing
    End If
    Workbooks("PERSONAL.XLSB").Activate
    Cells(i, 2) = pmat
Next i
End Sub

