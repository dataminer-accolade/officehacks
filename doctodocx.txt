Function EndsWith(str As String, suffix As String) As Boolean
     Dim suffixLen As Integer
     suffixLen = Len(suffix)
     EndsWith = (Right(Trim(UCase(str)), suffixLen) = UCase(suffix))
End Function

Sub GetFiles(folderPath, fileList)
    Dim fileSystem, directoryInfo
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set directoryInfo = fileSystem.GetFolder(folderPath)
    For Each file In directoryInfo.files
        fileList.Add file
    Next
    For Each subfolder In directoryInfo.SubFolders
        GetFiles subfolder.path, fileList
    Next
End Sub


Sub DocToDocx()
 Dim docxPath As String
 Dim path As String: path = "D:\docs"
 Dim wordApplication As New Word.Application
 Dim wordDocument As Word.Document
 Set fileList = CreateObject("System.Collections.ArrayList")
 
 GetFiles path, fileList
 
 For Each file In fileList
    docxPath = file.path & "x"
    If EndsWith(file.Name, "doc") And Dir(docxPath) = "" Then
      With wordApplication
        Set wordDocument = .Documents.Open(FileName:=file.path, AddToRecentFiles:=False, ReadOnly:=True, Visible:=False)
        With wordDocument
            .SaveAs FileName:=docxPath, FileFormat:=16
            .Close
        End With
      End With
    End If
 Next
 
 Set wordDocument = Nothing
 Set wordApplication = Nothing
End Sub
