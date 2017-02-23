Sub Main()
    Dim wordApp As Object
    Worksheets("Data").Activate

    Dim xlSheetCols
    Dim xlSheetRows

    xlSheetCols = ActiveSheet.UsedRange.Columns.Count
    xlSheetRows = ActiveSheet.UsedRange.Rows.Count

    ' ��Ҫ��δ��д����ʱʹ��
    If Cells(2, "A") = "" Then
        msg = MsgBox("���������ݺ���ʹ���Զ����ɹ���", vbOKOnly, "ע��")
        Exit Sub
    End If

    Dim strFolderName, strFolderPath As String
    strFolderName = Format(Now(), "����Ʊyyyy-mm-dd")
    strFolderRootPath = "D:\"
    strFilePath = strFolderRootPath & strFolderName & "\"

    If Dir(strFolderRootPath & strFolderName, vbDirectory) <> "" Then
        msg = MsgBox("Ŀ���ļ����Ѵ��ڣ��Ƿ�ɾ����", vbYesNo, "ע��")
        If msg = vbNo Then
            MsgBox "����ֹ���������ʹ�ã���ɾ��""" & strFolderName & """�ļ��к���������", vbOKOnly, "ע��"
            Exit Sub
        Else
            DeleteFolder(strFolderRootPath & strFolderName)
        End If
    End If

    MkDir strFolderRootPath & strFolderName
    
    Set wordApp = CreateObject("Word.Application")
    ' Word����Ŀɼ��ԣ�����Ҫʱ��ΪFalse
    wordApp.Visible = False

    Dim wdDocs As Documents
    Dim wdDoc As Document

    For i = 2 To xlSheetRows
        Application.StatusBar = "���ڴ����" & i - 1 & "����¼, ��" & xlSheetRows - 1 & "����¼"

        strTicketStationName = Cells(i, "B")
        strTicketId = Cells(i, "C")
        strTicketStartTime = Cells(i, "D")
        strTicketStopTime = Cells(i, "E")
        strFileName = Cells(i, "F")
        
        ' TODO: change here!
        Set wdDoc = wordApp.Documents.Add(Template:="C:\Users\Asterodeia\Desktop\����\���ڶ��ֹ���Ʊģ��2.dotx", NewTemplate:=False, DocumentType:=0)
        
        wdDoc.Activate
        With wordApp.Selection
            ' �ƶ����ĵ���ʼ
            .HomeKey Unit:=wdStory
            ' ����Ʊ���
            .NextField.Select
            .TypeText Text:=strTicketId
            ' �����������վ����
            .NextField.Select
            .TypeText Text:=strTicketStationName
            ' ɾ����β�����»���
            .EndKey Unit:=wdLine
            .MoveLeft Unit:=wdCharacter, Count:=(Len(strTicketStationName) * 2), Extend:=wdExtend
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
    Application.StatusBar = "Done��"
    Application.Wait(Now() + TimeValue("0:0:3"))
    Application.StatusBar = ""

End Sub

Sub DeleteFolder(ByVal path As String)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder(path)
    Set fso = Nothing
End Sub
