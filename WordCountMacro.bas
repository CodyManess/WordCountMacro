Sub WordCount()

    Dim docname As String
    Dim NumWords As Long
    Dim NumFiles As Integer
    Dim FD As FileDialog
    
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    
    With FD
        .Title = "Select the folder that contains the documents"
        
        If .Show = -1 Then
            PathName = .SelectedItems(1) & "\"
            
        Else
            MsgBox ("You did not select a folder.  Exiting")
            Exit Sub
        End If
    End With
    
    Set objExcel = CreateObject("Excel.Application")
    Set wb = objExcel.Workbooks.Add
    Set ws = wb.Sheets("Sheet1")
    
    NumWords = 0
    NumFiles = 0
    
    ws.Cells(NumFiles + 1, "A").Value = "Document Names"
    ws.Cells(NumFiles + 1, "B").Value = "Word Count"
    
    docname = Dir(PathName & "*.doc*")
    
    While docname <> ""
        NumFiles = NumFiles + 1
        Documents.Open FileName:=PathName & docname, Visible:=False
        Documents(docname).Activate
        
        NumWords = NumWords + ActiveDocument.BuiltInDocumentProperties("Number of Words").Value
        
        ws.Cells(NumFiles + 1, "A").Value = docname
        ws.Cells(NumFiles + 1, "B").Value = ActiveDocument.BuiltInDocumentProperties("Number of Words").Value
        
        Documents(docname).Close savechanges:=False
        
        docname = Dir
    Wend
    
    ws.Cells(NumFiles + 3, "A").Value = "Final Word Count: "
    ws.Cells(NumFiles + 3, "B").Value = NumWords
    wb.SaveAs PathName & "WordCounts"
    wb.Close
    
    MsgBox ("There are " & NumWords & " words in " & NumFiles & " documents.")

End Sub
