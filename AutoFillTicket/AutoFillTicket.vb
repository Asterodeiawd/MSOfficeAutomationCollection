Option Explicit

Sub Main()
    Dim wordApp As Object
    Worksheets("Data").Activate

    Dim xlSheetCols
    Dim xlSheetRows

'    xlSheetCols = ActiveSheet.UsedRange.Columns.Count
    xlSheetRows = ActiveSheet.UsedRange.Rows.Count
    
    ' 不要在未填写数据时使用
    Dim msg As Integer
    If Cells(2, "A") = "" Then
       msg = MsgBox("请输入数据后再使用自动生成功能", vbOKOnly, "注意")
       Exit Sub
    End If
    
    ' 读取目录参数
    Dim strFolderName, strFolderPath, strRequiredFolder, strFolderRootPath, strTemplatePath As String
    
    strFolderRootPath = Worksheets("Params").Range("B1").Value
    strFolderName = Worksheets("Params").Range("B2").Value
    
    strRequiredFolder = Worksheets("Params").Range("B3").Value
    If Application.Version <= 14 Then
        If strRequiredFolder <> "" Then
            strTemplatePath = strRequiredFolder
        Else
            strTemplatePath = Application.TemplatesPath
        End If
    Else
        strTemplatePath = Worksheets("Params").Range("B3").Value
    End If
    
    Dim strFilePath As String
    strFilePath = PathJoin(strFolderRootPath, strFolderName)

    If dir(strFilePath, vbDirectory) <> "" Then
        
        msg = MsgBox("目标文件夹已存在，是否删除并继续？", vbYesNo, "注意")
        If msg = vbNo Then
            MsgBox "已终止，如需继续使用，请删除""" & strFolderName & """文件夹后重新运行", vbOKOnly, "注意"
            Exit Sub
        Else
            DeleteFolder (strFilePath)
        End If
    End If
    
    MakeDir strFilePath
    
    Set wordApp = CreateObject("Word.Application")
    ' Word程序的可见性，不需要时改为False
    wordApp.Visible = False
    
    Dim wdDocs As Documents
    Dim wdDoc As Document
    
    Dim i As Integer
    Dim strTicketStationName, strTicketId, strTicketStartTime, strTicketStopTime, strFileName As String
    For i = 2 To xlSheetRows
        Application.StatusBar = "正在处理第" & i - 1 & "条记录, 共" & xlSheetRows - 1 & "条记录"
        
        strTicketStationName = Cells(i, "B").Value
        strTicketId = Cells(i, "C").Value
        strTicketStartTime = Cells(i, "D").Value
        strTicketStopTime = Cells(i, "E").Value
        strFileName = Cells(i, "F").Value
        
        ' TODO: change here!
        Set wdDoc = wordApp.Documents.Add(Template:=PathJoin(strTemplatePath, "变电第二种工作票模板.dotx"), _
            NewTemplate:=False, DocumentType:=0)
        
        wdDoc.Activate
        With wordApp.Selection
            ' 移动至文档起始
            .HomeKey Unit:=wdStory
            ' 工作票序号
            .NextField.Select
            .TypeText Text:=strFileName
            ' 工作任务：配电站名称
            .NextField.Select
            .TypeText Text:=strTicketStationName
            
            ' 删除结尾多余下划线, 中文字符宽度为2，英文字符宽度为1
            .EndKey Unit:=wdLine
            .MoveLeft Unit:=1, Count:=(LenB(StrConv(strTicketStationName, vbFromUnicode))), Extend:=1
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

        wdDoc.SaveAs Filename:=PathJoin(strFilePath, strFileName), FileFormat:= _
            wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
            :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
            :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
            SaveAsAOCELetter:=False
        
        If wordApp.Visible = True Then
            Application.Wait (Now() + TimeValue("0:0:3"))
        End If

        wdDoc.Close SaveChanges:=True

    Next
    
    wordApp.Quit
    Application.StatusBar = "Done！"
    Application.Wait (Now() + TimeValue("0:0:3"))
    Application.StatusBar = ""

End Sub

Private Function PathJoin(ByVal rootPath As String, ByVal dir As String, Optional ByVal separator As String = "\")

    If Right(rootPath, 1) <> separator Then
        PathJoin = rootPath & separator & dir
    Else
        PathJoin = rootPath & dir
    End If
    
End Function

Private Sub MakeDir(ByVal path As String)
    ' this needs Microsoft Scripting Runtime
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    MakeDirImpl path, fso

    Set fso = Nothing
End Sub

Private Function MakeDirImpl(ByVal path As String, ByRef fso As Object) As Boolean
  
    Dim parentFolder As String
    parentFolder = fso.GetParentFolderName(path)
    
    While Not fso.FolderExists(parentFolder)
        MakeDirImpl parentFolder, fso
    Wend
    
    fso.CreateFolder (path)
    MakeDirImpl = True
    
End Function

Sub DeleteFolder(ByVal path As String)
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder (path)

    Set fso = Nothing
End Sub
