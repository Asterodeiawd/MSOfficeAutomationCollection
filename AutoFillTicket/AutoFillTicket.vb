Option Explicit

Sub Main()
    Dim wordApp As Object
    Worksheets("Data").Activate

    Dim xlSheetCols
    Dim xlSheetRows

'    xlSheetCols = ActiveSheet.UsedRange.Columns.Count
    xlSheetRows = ActiveSheet.UsedRange.Rows.Count
    
    ' ��Ҫ��δ��д����ʱʹ��
    Dim msg As Integer
    If Cells(2, "A") = "" Then
       msg = MsgBox("���������ݺ���ʹ���Զ����ɹ���", vbOKOnly, "ע��")
       Exit Sub
    End If
    
    ' ��ȡĿ¼����
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
        
        msg = MsgBox("Ŀ���ļ����Ѵ��ڣ��Ƿ�ɾ����������", vbYesNo, "ע��")
        If msg = vbNo Then
            MsgBox "����ֹ���������ʹ�ã���ɾ��""" & strFolderName & """�ļ��к���������", vbOKOnly, "ע��"
            Exit Sub
        Else
            DeleteFolder (strFilePath)
        End If
    End If
    
    MakeDir strFilePath
    
    Set wordApp = CreateObject("Word.Application")
    ' Word����Ŀɼ��ԣ�����Ҫʱ��ΪFalse
    wordApp.Visible = False
    
    Dim wdDocs As Documents
    Dim wdDoc As Document
    
    Dim i As Integer
    Dim strTicketStationName, strTicketId, strTicketStartTime, strTicketStopTime, strFileName As String
    For i = 2 To xlSheetRows
        Application.StatusBar = "���ڴ����" & i - 1 & "����¼, ��" & xlSheetRows - 1 & "����¼"
        
        strTicketStationName = Cells(i, "B").Value
        strTicketId = Cells(i, "C").Value
        strTicketStartTime = Cells(i, "D").Value
        strTicketStopTime = Cells(i, "E").Value
        strFileName = Cells(i, "F").Value
        
        ' TODO: change here!
        Set wdDoc = wordApp.Documents.Add(Template:=PathJoin(strTemplatePath, "���ڶ��ֹ���Ʊģ��.dotx"), _
            NewTemplate:=False, DocumentType:=0)
        
        wdDoc.Activate
        With wordApp.Selection
            ' �ƶ����ĵ���ʼ
            .HomeKey Unit:=wdStory
            ' ����Ʊ���
            .NextField.Select
            .TypeText Text:=strFileName
            ' �����������վ����
            .NextField.Select
            .TypeText Text:=strTicketStationName
            
            ' ɾ����β�����»���, �����ַ����Ϊ2��Ӣ���ַ����Ϊ1
            .EndKey Unit:=wdLine
            .MoveLeft Unit:=1, Count:=(LenB(StrConv(strTicketStationName, vbFromUnicode))), Extend:=1
            .TypeBackspace
            
            ' �ƻ�����ʱ�䣺��ʼ
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

            ' �ƻ�����ʱ�䣺����
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
    Application.StatusBar = "Done��"
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
