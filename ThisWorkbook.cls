VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    
    Dim lastRow As Integer
    Dim currentMonth As Integer: currentMonth = month(Date)
    Dim currentDay As Integer: currentDay = day(Date)
    
    Worksheets("Main").Range("F2").Value = ""
    Worksheets("Main").Range("F3").Value = ""
    
    With ThisWorkbook.Worksheets("Main").ComboBox1
        .AddItem "Please Select a Month"
        .AddItem "01"
        .AddItem "02"
        .AddItem "03"
        .AddItem "04"
        .AddItem "05"
        .AddItem "06"
        .AddItem "07"
        .AddItem "08"
        .AddItem "09"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
        .ListIndex = currentMonth ' change back to 0
    End With
    
    With ThisWorkbook.Worksheets("Main").ComboBoxLog
        .AddItem "On"
        .AddItem "Off"
        .ListIndex = 1 ' change back to 0
    End With
    
    With ThisWorkbook.Worksheets("Main").ComboBox2
        .AddItem "Please Select a Day"
        .AddItem "01"
        .AddItem "02"
        .AddItem "03"
        .AddItem "04"
        .AddItem "05"
        .AddItem "06"
        .AddItem "07"
        .AddItem "08"
        .AddItem "09"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
        .AddItem "13"
        .AddItem "14"
        .AddItem "15"
        .AddItem "16"
        .AddItem "17"
        .AddItem "18"
        .AddItem "19"
        .AddItem "20"
        .AddItem "21"
        .AddItem "22"
        .AddItem "23"
        .AddItem "24"
        .AddItem "25"
        .AddItem "26"
        .AddItem "27"
        .AddItem "28"
        .AddItem "29"
        .AddItem "30"
        .AddItem "31"
        If currentDay > 8 Then
            .ListIndex = currentDay - 7
        Else
            .ListIndex = 0
        End If
    End With
 
     With ThisWorkbook.Worksheets("Main").ComboBox3
        .AddItem "Please Select a Year"
        .AddItem "2018"
        .AddItem "2017"
        .AddItem "2016"
        .AddItem "2015"
        .AddItem "2014"
        .AddItem "2013"
        .AddItem "2012"
        .AddItem "2011"
        .AddItem "2010"
        .ListIndex = 1 ' change back to 0
    End With
    
    With ThisWorkbook.Worksheets("Main").ComboBoxColors
        .AddItem "Yellow"
        .AddItem "Blue"
        .AddItem "Cyan/Light Blue"
        .AddItem "Light Green"
        .AddItem "Pink"
        .AddItem "Orange"
        .AddItem "Purple"
        .ListIndex = 1 ' change back to 0
    End With
    
    ' Clearing previous content
'    lastRow = Sheets(1).Cells(Sheets(1).Rows.Count, "A").End(xlUp).row
'    If lastRow <> 7 Then
'        ThisWorkbook.Worksheets(1).Range("A8:L" & lastRow).ClearContents
'    End If
    
End Sub
