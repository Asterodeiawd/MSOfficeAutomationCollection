Sub Main()
    Dim wordApp As Object
    Worksheets("Data").Activate

    Dim xlSheetCols
    Dim xlSheetRows

    xlSheetCols = ActiveSheet.UsedRange.Columns.Count
    xlSheetRows = ActiveSheet.UsedRange.Rows.Count

    ' 不要在未填写数据时使用
    If Cells(2, "A") = "" Then
        msg = MsgBox("请输入数据后再使用自动生成功能", vbOKOnly, "注意")
        Exit Sub
    End If

    Dim strFolderName, strFolderPath As String
    strFolderName = Format(Now(), "工作票yyyy-mm-dd")
    strFolderRootPath = "D:\"
    strFilePath = strFolderRootPath & strFolderName & "\"

    If Dir(strFolderRootPath & strFolderName, vbDirectory) <> "" Then
        msg = MsgBox("目标文件夹已存在，是否删除？", vbYesNo, "注意")
        If msg = vbNo Then
            MsgBox "已终止，如需继续使用，请删除""" & strFolderName & """文件夹后重新运行", vbOKOnly, "注意"
            Exit Sub
        Else
            DeleteFolder(strFolderRootPath & strFolderName)
        End If
    End If

    MkDir strFolderRootPath & strFolderName
    
    Set wordApp = CreateObject("Word.Application")
    ' Word程序的可见性，不需要时改为False
    wordApp.Visible = False

    Dim wdDocs As Documents
    Dim wdDoc As Document

    For i = 2 To xlSheetRows
        Application.StatusBar = "正在处理第" & i - 1 & "条记录, 共" & xlSheetRows - 1 & "条记录"

        strTicketStationName = Cells(i, "B")
        strTicketId = Cells(i, "C")
        strTicketStartTime = Cells(i, "D")
        strTicketStopTime = Cells(i, "E")
        strFileName = Cells(i, "F")
        
        ' TODO: change here!
        Set wdDoc = wordApp.Documents.Add(Template:="C:\Users\Asterodeia\Desktop\变电第二种工作票模板2.dotx", NewTemplate:=False, DocumentType:=0)
        
        wdDoc.Activate
        With wordApp.Selection
            ' 移动至文档起始
            .HomeKey Unit:=wdStory
            ' 工作票序号
            .NextField.Select
            .TypeText Text:=strTicketId
            ' 工作任务：配电站名称
            .NextField.Select
            .TypeText Text:=strTicketStationName
            ' 删除结尾多余下划线
            .EndKey Unit:=wdLine
            .MoveLeft Unit:=wdCharacter, Count:=(Len(strTicketStationName) * 2), Extend:=wdExtend
            .TypeBackspace

            ' 计划工作时间：起始
            .NextField.Select
            .TypeText Text:=Format(Year(strTicketStartTime), "0000")
            .NextField.Select
            .TypeText Text:=Format(Month(strTicketStartTime), "00")
            .NextField.Select
            .TypeText Text:=Format(Day(strTicketStartTime), "00")
            .NextField.Select
            .TypeText Text:=Format(Hour(strTicketStartTime), "00")
            .NextField.Select
            .TypeText Text:=Format(Minute(strTicketStartTime), "00")

            ' 计划工作时间：结束
            .NextField.Select
            .TypeText Text:=Format(Year(strTicketStopTime), "0000")
            .NextField.Select
            .TypeText Text:=Format(Month(strTicketStopTime), "00")
            .NextField.Select
            .TypeText Text:=Format(Day(strTicketStopTime), "00")
            .NextField.Select
            .TypeText Text:=Format(Hour(strTicketStopTime), "00")
            .NextField.Select
            .TypeText Text:=Format(Minute(strTicketStopTime), "00")
        End With

        wdDoc.SaveAs2 Filename:=strFilePath & strFileName, FileFormat:=
            wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
            :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False,
            SaveAsAOCELetter:=False, CompatibilityMode:=15

        If wordApp.Visible = True Then
            Application.Wait(Now() + TimeValue("0:0:3"))
        End If

        wdDoc.Close SaveChanges:=True

    Next

    wordApp.Quit
    Application.StatusBar = "Done！"
    Application.Wait(Now() + TimeValue("0:0:3"))
    Application.StatusBar = ""

End Sub

Sub DeleteFolder(ByVal path As String)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder(path)
    Set fso = Nothing
End Sub
