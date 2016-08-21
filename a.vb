Sub Main()
    Dim wdApp As Application
    Dim xlApp As Excel.Application

    
    Set wdApp = Application
    Set xlApp = CreateObject("Excel.Application")
'    xlApp.Visible = True
    xlApp.Visible = False

    
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet

    Set xlBook = xlApp.Workbooks.Open("E:\Users\Asterodeia\Documents\test.xls")

'    xlBook.Activate
    Set xlSheet = xlBook.Sheets(1)

    Dim xlSheetCols, xlSheetRows As Integer

    xlSheetCols = xlSheet.UsedRange.Columns.Count
    xlSheetRows = xlSheet.UsedRange.Rows.Count

    '   On Error Resume Next
    Dim str As String
    Dim strChoice(1 To 4) As String
    Dim maxAnswerLen, curlen As Integer
    Dim separater
    Dim ans
    '    For Row = 2 To 5
	'	 xls structure
	'
	'   |A	|B	|C	|D	|E	|F		|G		|H		|I		|J		|K		|L		|
	'						Question| <- Choices 1 to 4  -> |				|Answer |

	For Row = 2 To xlSheetRows
        Debug.Print ("Processing question " & (Row - 1) & ", " & (xlSheetRows - 1) & " in total")
        
'        xlSheet.Range("F" & Row).Activate

        str = xlApp.Cells(Row, "F")
        strChoice(1) = xlApp.Cells(Row, "G")
        strChoice(2) = xlApp.Cells(Row, "H")
        strChoice(3) = xlApp.Cells(Row, "I")
        strChoice(4) = xlApp.Cells(Row, "J")

        ans = Chr(Asc("A") + xlApp.Cells(Row, "L") - 1)
        wdApp.Selection.TypeText Text:=str
        Selection.MoveUp unit:=wdParagraph
        Selection.MoveDown unit:=wdParagraph, Extend:=wdExtend
        Selection.Style = ActiveDocument.Styles("±êÌâ 2")
        Selection.MoveDown unit:=wdParagraph
        wdApp.Selection.TypeParagraph


        maxAnswerLen = Len(strChoice(1))

        ' Get the max length of the 4 choices
        For i = 2 To 4
            curlen = Len(strChoice(i))
            If curlen > maxAnswerLen Then
                maxAnswerLen = curlen
            End If
        Next
        
        ' If the max length less equal than 7, choices will be displayed in one line
        ' separated by tab
        If maxAnswerLen <= 7 Then
            separater = vbTab

            Selection.ParagraphFormat.TabStops(CentimetersToPoints(1)).Position = _
            CentimetersToPoints(1.25)
            Selection.ParagraphFormat.TabStops(CentimetersToPoints(4.25)).Position = _
            CentimetersToPoints(4.5)
            Selection.ParagraphFormat.TabStops(CentimetersToPoints(7.5)).Position = _
            CentimetersToPoints(7.75)
            Selection.ParagraphFormat.TabStops(CentimetersToPoints(10.75)).Position = _
            CentimetersToPoints(11.25)
        Else
            separater = vbCrLf & vbTab
            Selection.ParagraphFormat.TabStops(CentimetersToPoints(1)).Position = _
            CentimetersToPoints(1.25)
        End If

        str = vbTab & "A) " & strChoice(1) & separater & "B) " & strChoice(2) & separater & "C) " & strChoice(3) & separater & "D) " & strChoice(4)
        wdApp.Selection.TypeText Text:=str
        wdApp.Selection.TypeParagraph
        wdApp.Selection.TypeText Text:="´ð°¸£º " & ans
        wdApp.Selection.TypeParagraph


    Next
    xlBook.Close
    xlApp.Quit

    Set xlBook = Nothing
    Set xlApp = Nothing
    Set xlSheet = Nothing
    
End Sub
