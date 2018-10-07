Attribute VB_Name = "Module3"
Option Explicit

Sub FilePicker()
 
    Dim wsName As String: wsName = "Main"
 
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
 
        If .SelectedItems.Count > 0 Then
            ThisWorkbook.Worksheets(wsName).Range("L5").Value = .SelectedItems(1)
        End If
        
    End With

End Sub

Sub Reset()

    Dim i As Long
    Dim lastRow As Integer
    Dim wsName As String: wsName = "Main"
    
    ' Clearing previous connections
    With ThisWorkbook
        For i = .Connections.Count To 1 Step -1
            .Connections(i).Delete
        Next i
    End With

    ' Clearing previous content
    With ThisWorkbook.Worksheets(wsName)
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        If lastRow <> 7 Then
            .Range("A8:P" & lastRow).ClearContents
            .Range("A8:P" & lastRow).Interior.ColorIndex = 0
        End If
    End With
    'ThisWorkbook.Worksheets(1).Range("Q8").Interior.Color = RGB(240, 155, 68)
    
End Sub 
