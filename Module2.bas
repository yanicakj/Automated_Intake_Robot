Attribute VB_Name = "Module2"
Option Explicit 

Sub PushToMasterTrackerSP()

    Dim wb As Workbook
    Dim MasterTracker As Workbook
    Dim lastLocalRow As Integer
    Dim lastLogRow As Integer
    Dim localRow As Integer
    Dim logCount As Integer: logCount = 0
    Dim row As Integer
    Dim currentWR As String
    Dim redValue As Integer: redValue = 0
    Dim greenValue As Integer: greenValue = 0
    Dim blueValue As Integer: blueValue = 0
    Dim colorValue As String
    Dim foundOnSheet As Boolean: foundOnSheet = False
    Dim username As String
    Dim password As String
    Dim usingSharePoint As Boolean: usingSharePoint = False
    Dim currentMonth As String
    Dim deliverableResponse As Boolean: deliverableResponse = False
    Dim wsName As String: wsName = "Main"
    Dim ws2Name As String: ws2Name = "Tables"
    Dim wSheet As Variant
    Dim rgFound As Range
    Dim stringArray As Variant
    Dim progRow As Integer: progRow = 1
    Dim i As Integer
    
    ' Returning username and password by reference
    If Not GetUsernameAndPassword(username, password, wsName) Then Exit Sub
        
    Application.StatusBar = "Pushing WR info to SharePoint"
    
    With ThisWorkbook.Worksheets(wsName)
        lastLocalRow = .Cells(.Rows.Count, "A").End(xlUp).row
        
        ' Cleaning right side
        .Range("M8:O" & lastLocalRow).ClearContents
        .Range("M8:O" & lastLocalRow).Interior.ColorIndex = 0
    End With

    ' Exiting program if there are no WRs to push
    If lastLocalRow = 7 Then
        MsgBox "No info to push"
        Application.StatusBar = ""
        Exit Sub
    End If
    
    ' Function to check overrides
    Set MasterTracker = OverrideHandler(usingSharePoint, wsName)
    If MasterTracker Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' Looping to find checked out master tracker
    For Each wb In Application.Workbooks
        If InStr(wb.Worksheets(1).Name, "Project Pipeline") > 0 Then
            Set MasterTracker = wb
            Exit For
        End If
    Next
    
    ' Finding last row on main tab of tool - TWICE?
    With ThisWorkbook.Worksheets(wsName)
        lastLocalRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    
    ' Removing previous intake cell color fills
    With ThisWorkbook.Worksheets("Log")
        lastLogRow = .Cells(.Rows.Count, "H").End(xlUp).row
        colorValue = .Range("H" & lastLogRow).Value
        GetColorValues redValue, greenValue, blueValue, colorValue, wsName
    End With
    RemovePreviousColorsAndFilters MasterTracker, redValue, greenValue, blueValue
    
    ' Setting new desired values
    With ThisWorkbook.Worksheets(wsName)
        colorValue = .ComboBoxColors.Value
        GetColorValues redValue, greenValue, blueValue, colorValue, wsName
    End With
    
    ' Checking for Dupes, creating temp tab
    ThisWorkbook.Sheets.Add.Name = "Temp"
    With ThisWorkbook.Worksheets("Temp")
        .Activate
        
        stringArray = Array("Please Wait. Checking for Dupes.", _
                                "Please Wait. Checking Intake Blanks Tab.", _
                                "Please Wait. Checking Intake No's Tab.", _
                                "Please Wait. Checking Not Testing Tab.", _
                                "Please Wait. Checking TBD Tab.")
        
        For i = LBound(stringArray) To UBound(stringArray)
            .Range("A" & progRow).ColumnWidth = 16
            .Range("K" & progRow).ColumnWidth = 15
            .Range("A" & progRow).Value = stringArray(i)
            .Range("K" & progRow).Value = "Finished"
            .Range("A" & progRow).HorizontalAlignment = xlCenter
            .Range("A" & progRow).VerticalAlignment = xlCenter
            .Range("K" & progRow).HorizontalAlignment = xlCenter
            .Range("K" & progRow).VerticalAlignment = xlCenter
            .Range("A" & progRow & ":K" & progRow).Borders.LineStyle = xlContinuous
            .Range("A" & progRow).WrapText = True
            .Rows(progRow).RowHeight = 60
            progRow = progRow + 2
        Next
        
    End With
    
    ' Iterating through WRs to check each one against Master Tracker
    ' Changed WS to 2 from 1 for all below
    Application.ScreenUpdating = False
    Application.StatusBar = "Checking for dupes against master tracker"
    
    With ThisWorkbook.Worksheets(wsName)
        For row = 8 To lastLocalRow
            
            ' Setup
            currentWR = Trim(CStr(.Range("A" & row).Value))
            
            ' Updating progress green bar
            CheckProgress lastLocalRow, row, "1", "none"
            
            ' Looping thru sheets
            For Each wSheet In MasterTracker.Worksheets
                With wSheet.UsedRange
                    Set rgFound = .Cells.Find(What:=currentWR, After:=.Range("A1"), LookIn:=xlValues, _
                        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False, SearchFormat:=False) ' .Activate
                End With
                
                If rgFound Is Nothing And Len(Trim(CStr(.Range("N" & row).Value))) = 0 Then
                    .Range("M" & row).Value = "Not on tracker"
                    .Range("M" & row).Interior.Color = RGB(255, 150, 150)
                ElseIf Not (rgFound Is Nothing) And Len(Trim(CStr(.Range("N" & row).Value))) = 0 Then
                    .Range("N" & row).NumberFormat = "@"
                    .Range("M" & row).Value = "On Tracker"
                    .Range("N" & row).Value = rgFound.Worksheet.Name
                    .Range("M" & row).Interior.ColorIndex = 0
                    Exit For
                ElseIf Not (rgFound Is Nothing) Then
                    .Range("N" & row).Value = .Range("N" & row).Value & vbCrLf & rgFound.Worksheet.Name
                End If
            Next
            .Range("N" & row).HorizontalAlignment = xlLeft
        Next row
    End With
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Distributing Work Requests"
    ThisWorkbook.Worksheets(wsName).Activate
    
    ' Distributing WRs
    With ThisWorkbook.Worksheets(wsName)
        For localRow = 8 To lastLocalRow
            
            If .Range("M" & localRow).Value = "Not on tracker" Then
            
                If Trim(CStr(.Range("J" & localRow).Value)) <> "In_Progress" And Trim(CStr(.Range("J" & localRow).Value)) <> "System_Route_Work" Then
                
                    If Trim(CStr(.Range("J" & localRow).Value)) <> "Withdrawn" And Trim(CStr(.Range("J" & localRow).Value)) <> "Closed" Then
                    
                        If Len(Trim(CStr(.Range("B" & localRow).Value))) > 0 Then
                            
                            If Trim(CStr(.Range("B" & localRow).Value)) = "Yes" Then
                            
                                ' Add to month tabs
                                MonthAdder localRow, MasterTracker, redValue, greenValue, blueValue, "Tool", wsName
                                
                            Else ' UAT = No
                            
                                ' Add to Intake No's
                                IntakeBlankOrNoAdder localRow, MasterTracker, "Intake No's", "Tool", wsName
                            
                            End If ' End of UAT Yes/No
                            
                        Else ' UAT is blank
                            
                            ' Add to Intake Blanks
                            IntakeBlankOrNoAdder localRow, MasterTracker, "Intake Blanks", "Tool", wsName
                            
                        End If ' End of Empty UAT Inv
                    
                    Else ' = Withdrawn
                    
                        ' Add to Intake No's
                        IntakeBlankOrNoAdder localRow, MasterTracker, "Intake No's", "Tool", wsName
                
                    End If
                    
                Else ' = In_Progress
                
                    ' Push In Progress to Blanks Tab
                    IntakeBlankOrNoAdder localRow, MasterTracker, "Intake Blanks", "Tool", wsName
                    
                End If ' End of In Progress
            
            End If ' End of Not on Tracker, Else do nothing
            
        Next ' End of Main For iterating through rows
    End With
    
    ' Resetting screen
    Application.StatusBar = "Scrubbing Master Tracker"
    ThisWorkbook.Worksheets("Temp").Activate
    Application.ScreenUpdating = True
    
    ' Re-checking WRs on tabs
    With ThisWorkbook.Worksheets(wsName)
    
        ' Iterate over Intake Blanks here
        Application.StatusBar = "Scrubbing Intake Blanks tab"
        IntakeBlanksAndNosChecker MasterTracker, username, password, redValue, greenValue, blueValue, "Intake Blanks", wsName, ws2Name
        Application.Wait Now + TimeValue("00:00:01")
        ThisWorkbook.Worksheets("Temp").Activate
        
        ' Iterate over Intake Nos here
        Application.StatusBar = "Scrubbing Intake No's tab"
        IntakeBlanksAndNosChecker MasterTracker, username, password, redValue, greenValue, blueValue, "Intake No's", wsName, ws2Name
        Application.Wait Now + TimeValue("00:00:01")
        ThisWorkbook.Worksheets("Temp").Activate
        
        ' Iterate over Not Testing here
        Application.StatusBar = "Scrubbing Not Testing tab"
        NotTestingChecker MasterTracker, username, password, wsName, ws2Name
        Application.Wait Now + TimeValue("00:00:01")
        ThisWorkbook.Worksheets("Temp").Activate
        
        ' Iterate over TBD tab here
        Application.StatusBar = "Scrubbing TBD tab"
        TBDChecker MasterTracker, username, password, wsName, ws2Name
        Application.Wait Now + TimeValue("00:00:01")
        ThisWorkbook.Worksheets("Temp").Activate
    
    End With
    
    ' Cleaning up temp sheet
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Temp").Delete
    Application.DisplayAlerts = True
    
'        ' Calling function to update deliverable dates on current month tab here
'        deliverableResponse = UpdateDeliverableDates(MasterTracker, username, password)
'
'        If deliverableResponse = False Then
'            ThisWorkbook.Worksheets("Main").Range("F2").Activate
'            Application.StatusBar = ""
'            ThisWorkbook.Worksheets(1).Range("A1:H1").Interior.ColorIndex = 0
'            If usingSharePoint = True Then
'                If MasterTracker.CanCheckIn = True Then
'                    MasterTracker.CheckIn SaveChanges:=False
'                End If
'            End If
'            MsgBox "Unable to connect to database with entered credentials."
'            Exit Sub
'        End If
    
    ' Cleanup
    With ThisWorkbook.Worksheets(wsName)
    
        ' Checking workbook back in
        Application.StatusBar = ""
        .Activate
        If usingSharePoint = True Then
            If MasterTracker.CanCheckIn = True Then
                'MasterTracker.CheckIn
                MsgBox "Master Tracker has been checked in."
            Else
                MsgBox "This file cannot be checked in at this time. Please try again later."
            End If
        Else
            MsgBox "Information has been pushed!"
        End If
        
        ' Recoloring green bar on top of sheet
        .Range("A1:H1").Interior.ColorIndex = 0
        
        ' Call function to generate email to Mani & Dan
        FinalEmailGenerator CStr(colorValue)
        
        ' update log
        If .ComboBoxLog.Value = "On" Then
        
            For row = 8 To lastLocalRow
                If .Range("M" & row).Value = "Not on tracker" Then
                    logCount = logCount + 1
                End If
            Next
        
            With ThisWorkbook.Worksheets("Log")
                lastLogRow = .Cells(.Rows.Count, "E").End(xlUp).row + 1
                .Range("E" & lastLogRow).Value = Date
                .Range("E" & lastLogRow).NumberFormat = "m/d/yyyy"
                .Range("F" & lastLogRow).Value = ThisWorkbook.Worksheets(1).Range("E8").Value
                .Range("G" & lastLogRow).Value = logCount
                .Range("H" & lastLogRow).Value = colorValue
            End With
        End If
    End With
    
    Exit Sub

' Error handler for DB2 connection
ConnectionHandler:
    ThisWorkbook.Worksheets(wsName).Range("F2").Activate
    MsgBox "Unable to connect to database. Please make sure credentials are correct!"
    Exit Sub

End Sub

Sub CheckProgress(lastLocalRow As Integer, row As Integer, checkRow As String, quitInd As String)
    
    If quitInd = "start" Then
        If row > (lastLocalRow / 2) Then
            Exit Sub
        End If
    ElseIf quitInd = "end" Then
        If row < (lastLocalRow / 2) Then
            Exit Sub
        End If
    End If
    
    ' Checking status
    With ThisWorkbook.Worksheets("Temp")
        If Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 10 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 20 Then
            MakeGreen "B" & checkRow
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 20 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 30 Then
            MakeGreen "C" & checkRow
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 30 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 40 Then
            MakeGreen "D" & checkRow
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 40 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 50 Then
            MakeGreen "E" & checkRow
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 50 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 60 Then
            MakeGreen "F" & checkRow
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 60 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 70 Then
            MakeGreen "G" & checkRow
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 70 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 80 Then
            MakeGreen "H" & checkRow
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 80 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 90 Then
            MakeGreen "I" & checkRow
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 90 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 100 Then
            MakeGreen "J" & checkRow
        End If
    End With

End Sub

Sub MakeGreen(rangeString As String)

    With ThisWorkbook.Worksheets("Temp")
        ' removed screenupdating lines
        .Range(rangeString).Interior.Color = RGB(0, 200, 0)
    End With

End Sub

Sub GetColorValues(ByRef redValue As Integer, ByRef greenValue As Integer, ByRef blueValue As Integer, colorValue As String, wsName As String)
    
    Select Case colorValue
        Case "Yellow"
            redValue = 255
            greenValue = 255
            blueValue = 0
        Case "Blue"
            redValue = 0
            greenValue = 125
            blueValue = 255
        Case "Cyan/Light Blue"
            redValue = 0
            greenValue = 255
            blueValue = 255
        Case "Light Green"
            redValue = 145
            greenValue = 255
            blueValue = 0
        Case "Pink"
            redValue = 255
            greenValue = 0
            blueValue = 255
        Case "Orange"
            redValue = 255
            greenValue = 125
            blueValue = 0
        Case "Purple"
            redValue = 125
            greenValue = 0
            blueValue = 255
        Case Else
            redValue = 255
            greenValue = 255
            blueValue = 0
    End Select
    

End Sub

Public Function GetUsernameAndPassword(ByRef username As String, ByRef password As String, wsName As String) As Boolean

    GetUsernameAndPassword = True

    ' checking user fields
    With ThisWorkbook.Worksheets(wsName)
        If Len(Trim(CStr(.Range("F2").Value))) = 0 Then
            MsgBox "Please enter your LAN ID!"
            .Range("F2").Activate
            GetUsernameAndPassword = False
            Exit Function
        Else
            username = Trim(CStr(.Range("F2").Value))
        End If
        
        If Len(Trim(CStr(.Range("F3").Value))) = 0 Then
            MsgBox "Please enter your LAN Password!"
            .Range("F3").Activate
            GetUsernameAndPassword = False
            Exit Function
        Else
            password = Trim(CStr(.Range("F3").Value))
        End If
    End With

End Function

Function OverrideHandler(ByRef usingSharePoint As Boolean, wsName As String) As Workbook

    Dim SharePointFile As String
    Dim tempWB As String
    
    Set OverrideHandler = Nothing

    ' Checking for overrides
    With ThisWorkbook.Worksheets(wsName)
        
        ' If both overrides are filled
        If Len(Trim(CStr(.Range("L4").Value))) > 0 And Len(Trim(CStr(.Range("L5").Value))) > 0 Then
            .Range("L4").Activate
            .Range("L4").Hyperlinks.Delete
            .Range("L4:N4").Borders.LineStyle = xlContinuous
            .Range("L4:N4").Borders.Color = RGB(0, 112, 192)
            
            MsgBox "Can't have both override fields populated!"

        ' If only SP override is filled
        ElseIf Len(Trim(CStr(.Range("L4").Value))) > 0 Then
            .Range("L4").Hyperlinks.Delete
            .Range("L4:N4").Borders.LineStyle = xlContinuous
            .Range("L4:N4").Borders.Color = RGB(0, 112, 192)
            
            usingSharePoint = True
            SharePointFile = Trim(CStr(.Range("L4").Value))
            
            If Workbooks.CanCheckOut(SharePointFile) = True Then
                Workbooks.CheckOut (SharePointFile)
                Workbooks.Open (SharePointFile)
                Application.Wait (Now + TimeValue("0:00:01"))
                tempWB = Replace(returnFileName(SharePointFile, "/"), "%20", " ")
                Set OverrideHandler = Workbooks(tempWB)
                
            Else
                MsgBox "Master Tracker can't be checked out at this time. Please make sure it is not open on your computer.", vbInformation
                Application.StatusBar = ""
            End If
        
        ' If only file override is filled
        ElseIf Len(Trim(CStr(.Range("L5").Value))) > 0 Then
            
            Set OverrideHandler = AttachMasterTracker(wsName, Trim(CStr(.Range("L5").Value)))
            
        ' If nothing is filled
        Else
            usingSharePoint = True
            SharePointFile = "..."
            
            If Workbooks.CanCheckOut(SharePointFile) = True Then
                Workbooks.CheckOut (SharePointFile)
                Workbooks.Open (SharePointFile)
                Application.Wait (Now + TimeValue("0:00:01"))
                
            Else
                MsgBox "Master Tracker can't be checked out at this time. Please make sure it is not open on your computer.", vbInformation
                Application.StatusBar = ""
            End If
        End If
    End With
    
    Exit Function
    
' Error handler if can't find override file
DocumentHandler:
    ThisWorkbook.Activate
    MsgBox "Please make sure filename is correct!"
    OverrideHandler = False

End Function

Sub MonthAdder(localRow As Integer, MasterTracker As Workbook, redValue As Integer, greenValue As Integer, blueValue As Integer, indicator As String, wsName As String)

    Dim monthTab As String: monthTab = ""
    Dim monthString As String
    Dim monthNum As Integer: monthNum = 0
    Dim lastMasterRow As Integer
    Dim releaseYear As String: releaseYear = ""
    Dim ws As Variant
    Dim wsNew As Worksheet
    Dim currentMonth As Integer
    Dim currentYear As Integer
    
    ' finding release date for wr
    If indicator = "Tool" Then
        With ThisWorkbook.Worksheets(wsName)
            If IsDate(.Range("I" & localRow).Value) And Len(Trim(CStr(.Range("I" & localRow).Value))) > 0 Then
                monthString = Format(.Range("I" & localRow).Value, "mmm")
                releaseYear = year(.Range("I" & localRow).Value)
                monthNum = month(.Range("I" & localRow).Value)
            Else
                monthString = "TBD"
            End If
        End With
        
    ElseIf indicator = "Blanks" Or indicator = "No's" Then
        With MasterTracker.Worksheets("Intake " & indicator)
            If IsDate(.Range("D" & localRow).Value) And Len(Trim(CStr(.Range("D" & localRow).Value))) > 0 Then
                monthString = Format(.Range("D" & localRow).Value, "mmm")
                releaseYear = year(.Range("D" & localRow).Value)
                monthNum = month(.Range("D" & localRow).Value)
            Else
                monthString = "TBD"
            End If
        End With
    End If
    
    ' check if month is out-dated
    If monthNum <> 0 Then
        currentMonth = month(Date)
        currentYear = year(Date)
        
        If currentYear > releaseYear Then
            currentMonth = currentMonth + 12
        ElseIf currentYear < releaseYear Then
            currentMonth = currentMonth - 12
        End If
        
        If currentMonth > monthNum Then
            monthString = "TBD"
            releaseYear = ""
        
            If Abs(currentMonth) - Abs(monthNum) = 1 Then
                monthString = Format(Date, "mmm")
                releaseYear = year(Date)
            End If
        End If
    End If
    
    For Each ws In MasterTracker.Worksheets
        If InStr(ws.Name, monthString) > 0 And InStr(ws.Name, releaseYear) Then
            monthTab = ws.Name
            Exit For
        End If
    Next
    
    If Len(monthTab) = 0 Then ' took out:  And CInt(releaseYear) <= year(Date) + 3
        With MasterTracker
            Set wsNew = .Sheets.Add(After:= _
                        .Sheets(.Sheets.Count - 2))
            wsNew.Name = monthString & " " & releaseYear
            monthTab = wsNew.Name
            
            ' headers
            With wsNew
                .Range("A2").Value = "UAT-COE Lead"
                .Range("B2").Value = "Project Name"
                .Range("C2").Value = "WR"
                .Range("D2").Value = "Release Date"
                .Range("E2").Value = "UAT-COE SME for UAT-COE Sign off"
                .Range("F2").Value = "Status"
                .Range("A1:F1").Merge
                .Range("A1").RowHeight = 18
                .Range("A1").Font.Size = 14
                .Range("A1").Font.Bold = True
                .Range("A2").ColumnWidth = 16.56
                .Range("B2").ColumnWidth = 53.78
                .Range("C2").ColumnWidth = 16.33
                .Range("D2").ColumnWidth = 18.56
                .Range("E2").ColumnWidth = 29.78
                .Range("F2").ColumnWidth = 40.89
                .Range("A1").Value = "UAT-COE Deliverables Review and Sign off (Test Plan, Testable Requirements Doc, Test SN/Cases, Test Results)"
                .Range("A2:F2").Interior.Color = RGB(230, 184, 183)
                .Range("A1:F1").Interior.Color = RGB(0, 176, 240)
                .Range("A1").Font.Bold = True
                .Range("A1:F1").WrapText = True
                .Range("A2:F2").WrapText = True
                .Range("A1:F2").Borders.LineStyle = xlContinuous
                .Range("A1:F2").HorizontalAlignment = xlCenter
            End With
        End With
    End If
    
    ' Finding last used row on master tracker to use as a marker when updating tracker
    With MasterTracker.Worksheets(monthTab)
        lastMasterRow = .Cells(.Rows.Count, "C").End(xlUp).row + 1
    End With
    
    ' Updating master tracker programmatically
    If indicator = "Tool" Then
        With MasterTracker.Worksheets(monthTab)
            ' Formatting
            FormatRow MasterTracker.Worksheets(monthTab), lastMasterRow, redValue, greenValue, blueValue
        
            .Range("A" & lastMasterRow).Value = "TBD"
            .Range("B" & lastMasterRow).Value = CStr(ThisWorkbook.Worksheets(wsName).Range("F" & localRow).Value)
            .Range("C" & lastMasterRow).Value = CStr(ThisWorkbook.Worksheets(wsName).Range("A" & localRow).Value)
            .Range("D" & lastMasterRow).Value = CStr(ThisWorkbook.Worksheets(wsName).Range("I" & localRow).Value)
            .Range("F" & lastMasterRow).Value = "Initiation - " & Trim(CStr(ThisWorkbook.Worksheets(wsName).Range("J" & localRow).Value))
        End With
        ThisWorkbook.Worksheets(wsName).Range("O" & localRow).Value = "Pushed to " & monthTab
    
    ElseIf indicator = "Blanks" Or indicator = "No's" Then
        With MasterTracker.Worksheets(monthTab)
            ' Formatting
            FormatRow MasterTracker.Worksheets(monthTab), lastMasterRow, redValue, greenValue, blueValue
        
            .Range("A" & lastMasterRow).Value = "TBD"
            .Range("B" & lastMasterRow).Value = CStr(MasterTracker.Worksheets("Intake " & indicator).Range("B" & localRow).Value)
            .Range("C" & lastMasterRow).Value = CStr(MasterTracker.Worksheets("Intake " & indicator).Range("C" & localRow).Value)
            .Range("D" & lastMasterRow).Value = CStr(MasterTracker.Worksheets("Intake " & indicator).Range("D" & localRow).Value)
            .Range("F" & lastMasterRow).Value = "Initiation - " & Trim(CStr(MasterTracker.Worksheets("Intake " & indicator).Range("F" & localRow).Value))
            
            ' Delete row
            MasterTracker.Worksheets("Intake " & indicator).Rows(localRow).EntireRow.Delete
        End With
    End If

End Sub

Sub FormatRow(targetWS As Worksheet, row As Integer, redValue As Integer, greenValue As Integer, blueValue As Integer)

    With targetWS
        .Range("A" & row & ":F" & row).Interior.Color = RGB(redValue, greenValue, blueValue)
        .Range("A" & row & ":F" & row).Borders.LineStyle = xlContinuous
        .Range("A" & row & ":C" & row).NumberFormat = "@"
        .Range("E" & row & ":F" & row).NumberFormat = "@"
        .Range("D" & row).NumberFormat = "m/d/yyyy"
        .Range("A" & row).HorizontalAlignment = xlCenter
        .Range("B" & row).HorizontalAlignment = xlLeft
        .Range("C" & row & ":F" & row).HorizontalAlignment = xlCenter
        .Range("A" & row & ":F" & row).VerticalAlignment = xlCenter
        
        If redValue = 0 And greenValue = 0 And blueValue = 0 Then
            .Range("A" & row & ":F" & row).Interior.ColorIndex = 0
        End If
    End With

End Sub

Sub IntakeBlankOrNoAdder(localRow As Integer, MasterTracker As Workbook, monthTab As String, indicator As String, wsName As String)
                
    Dim lastMasterRow As Integer
                
    ' Finding last used row on master tracker to use as a marker when updating tracker
    With MasterTracker.Worksheets(monthTab)
        lastMasterRow = .Cells(.Rows.Count, "C").End(xlUp).row + 1
    End With
    
    ' Updating master tracker programmetically
    If indicator = "Tool" Then
        With MasterTracker.Worksheets(monthTab)
            .Range("B" & lastMasterRow).Value = CStr(ThisWorkbook.Worksheets(wsName).Range("F" & localRow).Value)
            .Range("C" & lastMasterRow).Value = CStr(ThisWorkbook.Worksheets(wsName).Range("A" & localRow).Value)
            .Range("D" & lastMasterRow).Value = CStr(ThisWorkbook.Worksheets(wsName).Range("I" & localRow).Value)
            .Range("F" & lastMasterRow).Value = CStr(ThisWorkbook.Worksheets(wsName).Range("J" & localRow).Value)
            .Range("G" & lastMasterRow).Value = CStr(ThisWorkbook.Worksheets(wsName).Range("B" & localRow).Value)
            .Range("H" & lastMasterRow).Value = CStr(ThisWorkbook.Worksheets(wsName).Range("D" & localRow).Value)
            .Range("D" & lastMasterRow).NumberFormat = "m/d/yyyy"
            .Range("H" & lastMasterRow).NumberFormat = "m/d/yyyy"
            .Range("A" & lastMasterRow & ":H" & lastMasterRow).Borders.LineStyle = xlContinuous
        End With
        ThisWorkbook.Worksheets(wsName).Range("O" & localRow).Value = "Pushed to " & monthTab
    
    ElseIf indicator = "Blanks" Then
        With MasterTracker.Worksheets(monthTab)
            .Range("B" & lastMasterRow).Value = CStr(MasterTracker.Worksheets("Intake Blanks").Range("B" & localRow).Value)
            .Range("C" & lastMasterRow).Value = CStr(MasterTracker.Worksheets("Intake Blanks").Range("C" & localRow).Value)
            .Range("D" & lastMasterRow).Value = CStr(MasterTracker.Worksheets("Intake Blanks").Range("D" & localRow).Value)
            .Range("F" & lastMasterRow).Value = CStr(MasterTracker.Worksheets("Intake Blanks").Range("F" & localRow).Value)
            .Range("G" & lastMasterRow).Value = CStr(MasterTracker.Worksheets("Intake Blanks").Range("G" & localRow).Value)
            .Range("H" & lastMasterRow).Value = CStr(MasterTracker.Worksheets("Intake Blanks").Range("H" & localRow).Value)
            .Range("D" & lastMasterRow).NumberFormat = "m/d/yyyy"
            .Range("H" & lastMasterRow).NumberFormat = "m/d/yyyy"
            .Range("A" & lastMasterRow & ":H" & lastMasterRow).Borders.LineStyle = xlContinuous
            MasterTracker.Worksheets("Intake Blanks").Rows(localRow).EntireRow.Delete
        End With
    End If

End Sub


Sub IntakeBlanksAndNosChecker(MasterTracker As Workbook, username As String, password As String, redValue As Integer, greenValue As Integer, blueValue As Integer, tabChecking As String, wsName As String, ws2Name As String)

    Dim lastMTRow As Integer
    Dim row As Integer
    Dim conn As Object 'Variable for ADODB.Connection object
    Dim rs As Object 'Variable for ADODB.Recordset object
    Dim workRequest As String
    Dim extra As Integer
    Dim localCounter As Integer
    Dim secondSheetSpot As Integer
    Dim checkRow As String

    If tabChecking = "Intake Blanks" Then
        checkRow = "3"
    Else
        checkRow = "5"
    End If

    ' Instantiating connection objects
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    With MasterTracker.Worksheets(tabChecking)

        lastMTRow = .Cells(.Rows.Count, "C").End(xlUp).row
        
        On Error GoTo ConnectionHandler3
        conn.ConnectionString = "Provider=IBMDADB2.1;UID=" & username & ";PWD=" & password & ";Data Source=...;ProviderType=OLEDB"
        conn.Open
        
        ' Iterating through WRs
        On Error GoTo 0
        For row = 3 To lastMTRow
        
            ' update progress bar
            CheckProgress lastMTRow, row, checkRow, "start"
        
            workRequest = Trim(CStr(.Range("C" & row).Value))
            rs.Open "Select v..." & _
                    "... from ... as t1 inner join ... as t2 on t1... = t2... where t1.... = '" & workRequest & "' with ur", conn
            If Not (rs.BOF Or rs.EOF) Then
                .Range("D" & row).Value = rs.Fields(0).Value
                .Range("F" & row).Value = rs.Fields(1).Value
                .Range("G" & row).Value = rs.Fields(2).Value
                .Range("D" & row).NumberFormat = "m/d/yyyy"
            End If
            rs.Close
        Next
    
        ' Closing connection
        conn.Close
        
        ' Checking for moveable WRs
        For row = 3 To lastMTRow
        
            ' update progress bar
            CheckProgress lastMTRow, row, checkRow, "end"
        
            If Trim(CStr(.Range("F" & row).Value)) <> "In_Progress" And Trim(CStr(.Range("F" & row).Value)) <> "System_Route_Work" Then ' Add system_route_work
                
                If Trim(CStr(.Range("F" & row).Value)) <> "Withdrawn" And Trim(CStr(.Range("F" & row).Value)) <> "Closed" Then
                
                    If Len(Trim(CStr(.Range("G" & row).Value))) > 0 Then
                        
                        If Trim(CStr(.Range("G" & row).Value)) = "Yes" Then
                            
                            ' Add to months
                            If tabChecking = "Intake Blanks" Then
                                MonthAdder row, MasterTracker, redValue, greenValue, blueValue, "Blanks", wsName
                            ElseIf tabChecking = "Intake No's" Then
                                MonthAdder row, MasterTracker, redValue, greenValue, blueValue, "No's", wsName
                            End If
                            row = row - 1
                        
                        Else ' UAT = No
                        
                            ' Add to Intake No's
                            If tabChecking = "Intake Blanks" Then
                                IntakeBlankOrNoAdder row, MasterTracker, "Intake No's", "Blanks", wsName
                                row = row - 1
                            End If
                            
                        End If ' End of UAT yes/no

                    End If ' End of UAT blank
                
                Else ' Is Withdrawn or Closed
                    
                    ' Add to Intake No's
                    If tabChecking = "Intake Blanks" Then
                        IntakeBlankOrNoAdder row, MasterTracker, "Intake No's", "Blanks", wsName
                        row = row - 1
                    End If
                
                End If ' End of Withdrawn check
            
            End If ' End of In Progress check
        Next
        
    End With
    
    Exit Sub
    
ConnectionHandler3:
    ThisWorkbook.Activate
    MsgBox "Unable to connect to database. Was checking Master Tracker."

End Sub


Sub NotTestingChecker(MasterTracker As Workbook, username As String, password As String, wsName As String, ws2Name As String)

    Dim conn As Object 'Variable for ADODB.Connection object
    Dim rs As Object 'Variable for ADODB.Recordset object
    Dim yesWorkRequests As Collection: Set yesWorkRequests = New Collection
    Dim lastMTRow As Integer
    Dim workRequest As String
    Dim tempState As String
    Dim row As Integer
    
    ' Instantiating connection objects
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ConnectionHandler5
    conn.ConnectionString = "Provider=IBMDADB2.1;UID=" & username & ";PWD=" & password & ";Data Source=...;ProviderType=OLEDB"
    conn.Open
    
    On Error GoTo 0
    With MasterTracker.Worksheets("Not Testing")
        
        lastMTRow = .Cells(.Rows.Count, "C").End(xlUp).row
        
        For row = 2 To lastMTRow
        
            ' update progress bar
            CheckProgress lastMTRow, row, "7", "none"
        
            workRequest = Trim(CStr(.Range("C" & row).Value))
            rs.Open "Select ... " & _
                    "from ... as t1 inner join ... as t2 on t1.... = t2.... where t1.... = '" & workRequest & "' with ur", conn
            If Not (rs.BOF Or rs.EOF) Then
                tempState = rs.Fields(1).Value
                If tempState <> "Withdrawn" And tempState <> "Closed" Then
                    If rs.Fields(0).Value = "Yes" Then
                        If Len(CStr(rs.Fields(2).Value)) > 0 Then
                            If CDate(rs.Fields(2).Value) >= Date Then
                                yesWorkRequests.Add workRequest
                            End If
                        End If
                    End If
                End If
            End If
            rs.Close
        Next
    
    End With
    
    If yesWorkRequests.Count > 0 Then
        EmailGenerator yesWorkRequests, "notTesting", username, password, wsName
    End If
    
    Exit Sub
    
ConnectionHandler5:
    ThisWorkbook.Activate
    MsgBox "Unable to connect to database. Was checking Master Tracker."

End Sub

Sub TBDChecker(MasterTracker As Workbook, username As String, password As String, wsName As String, ws2Name As String)

    Dim conn As Object 'Variable for ADODB.Connection object
    Dim rs As Object 'Variable for ADODB.Recordset object
    Dim flaggedTBDs As Collection: Set flaggedTBDs = New Collection
    Dim lastMTRow As Integer
    Dim workRequest As String
    Dim tempState As String
    Dim row As Integer
    
    ' Instantiating connection objects
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ConnectionHandler6
    conn.ConnectionString = "Provider=IBMDADB2.1;UID=" & username & ";PWD=" & password & ";Data Source=...;ProviderType=OLEDB"
    conn.Open
    
    On Error GoTo 0
    With MasterTracker.Worksheets("TBD")
        
        lastMTRow = .Cells(.Rows.Count, "C").End(xlUp).row
        
        For row = 3 To lastMTRow
        
            ' update progress bar
            CheckProgress lastMTRow, row, "9", "none"
        
            workRequest = Trim(CStr(.Range("C" & row).Value))
            rs.Open "Select ..." & _
                    "from ... as t1 inner join ... as t2 on t1.... = t2.... where t1.... = '" & workRequest & "' with ur", conn
            If Not (rs.BOF Or rs.EOF) Then
                tempState = rs.Fields(1).Value
                
                If tempState = "Withdrawn" Or tempState = "Closed" Or tempState = "Released" Then
                    
                    ' move to withdrawn tab & delete
                    AddTBDtoWithdrawnTab row, workRequest, rs.Fields(0).Value, rs.Fields(1).Value, rs.Fields(2).Value, MasterTracker
                    .Rows(row).EntireRow.Delete
                    row = row - 1
                    
                ElseIf tempState <> "Rejected" And tempState <> "In_Progress" And tempState <> "System_Route_Work" And tempState <> "Replan" And tempState <> "Pending_CCB_Review" Then
                    
                    If IsNull(rs.Fields(2).Value) = False Then
                    
                        If CDate(rs.Fields(2).Value) > Date Then
                            
                            If rs.Fields(0).Value = "Yes" Then
                                
                                    flaggedTBDs.Add workRequest
                            
                            End If
                        
                        End If
                    
                    End If
                    
                End If
                
            End If
            rs.Close
        Next
    End With
    
    If flaggedTBDs.Count > 0 Then
        EmailGenerator flaggedTBDs, "tbdTab", username, password, wsName
    End If
    
    Exit Sub
    
ConnectionHandler6:
    ThisWorkbook.Activate
    MsgBox "Unable to connect to database. Was checking Master Tracker."

End Sub

Sub AddTBDtoWithdrawnTab(row As Integer, workRequest As String, involvement As String, state As String, releaseDate As String, MasterTracker As Workbook)

    Dim withdrawnTab As Worksheet
    Dim lastWithdrawnRow As Integer
    Dim ws As Variant
    
    ' finding withdrawn tab
    For Each ws In MasterTracker.Worksheets
        If InStr(LCase(ws.Name), "withdrawn") Then
            Set withdrawnTab = ws
            Exit For
        End If
    Next
    
    With withdrawnTab
        lastWithdrawnRow = .Cells(.Rows.Count, "C").End(xlUp).row + 1
        .Range("A" & lastWithdrawnRow).Value = MasterTracker.Worksheets("TBD").Range("A" & row).Value
        .Range("B" & lastWithdrawnRow).Value = MasterTracker.Worksheets("TBD").Range("B" & row).Value
        .Range("C" & lastWithdrawnRow).Value = workRequest
        .Range("D" & lastWithdrawnRow).Value = releaseDate
        .Range("E" & lastWithdrawnRow).Value = MasterTracker.Worksheets("TBD").Range("E" & row).Value
        .Range("F" & lastWithdrawnRow).Value = state
        .Range("G" & lastWithdrawnRow).Value = involvement
        
        FormatRow withdrawnTab, lastWithdrawnRow, 0, 0, 0
    End With

End Sub

Public Function BookOpen(strBookName As String) As Boolean

    Dim oBk As Workbook
    On Error Resume Next
    Set oBk = Workbooks(strBookName)
    On Error GoTo 0
    If oBk Is Nothing Then
        BookOpen = False
    Else
        BookOpen = True
    End If

End Function

Sub FinalEmailGenerator(colorValue As String)

    Dim olApp As Outlook.Application
    Dim newMail As Outlook.MailItem
    Dim emailBody As String
    Dim emailSig As String
    Dim fontColor As String

    Set olApp = CreateObject("Outlook.Application")  'Another option: Set olApp = New Outlook.Application
    Set newMail = olApp.CreateItem(olMailItem)
    
    Select Case colorValue
        Case "Yellow"
            fontColor = "#dade52"
        Case "Blue"
            fontColor = "#0016ff"
        Case "Cyan/Light Blue"
            fontColor = "#45e4db"
        Case "Light Green"
            fontColor = "#45e445"
        Case "Pink"
            fontColor = "#ff00ba"
        Case "Orange"
            fontColor = "#ffb300"
        Case "Purple"
            fontColor = "#a400ff"
        Case Else
            fontColor = "#000000"
    End Select

    emailBody = "<p style=""font-family:calibri;font-size:11pt"">Hi Guys,</p><p style=""font-family:calibri;font-size:11pt"">Please find attached the updated Master Tracker as of " & Date & " with intake WRs from this past week added in <b><span style=""color:" & fontColor & """>" & colorValue & ".</span></b></p>"
    emailSig = "<p style=""font-family:calibri;font-size:11pt"">Thank you,<br><br>Jacob Yanicak<br>UAT-COE Analyst<br>Horizon Blue Cross Blue Shield of NJ<br>....com<br>973-632-4337</p>"
     
    emailBody = emailBody & emailSig
     
    With newMail
        .To = "....com; ....com"
        .Subject = "Updated Master Tracker " & CStr(Date)
        .SentOnBehalfOfName = "....com"
        .HTMLBody = emailBody
        .Save
        .Close olPromptForSave
    End With
    
End Sub

Function UpdateDeliverableDates(MasterTracker As Workbook, username As String, password As String) As Boolean

    Dim currentMonth As String
    Dim startRow As Integer
    Dim lastRow As Integer
    Dim currentWR As String
    Dim conn As Object
    Dim rs As Object
    
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Getting month to check deliverables
    If day(Date) < 25 Then
        currentMonth = Format(Date, "mmm")
    Else
        currentMonth = month(DateAdd("m", 1, Date))
    End If
    
    ' Looping through master tracker to find correct tab
    
    On Error GoTo ConnectionHandler
    conn.Connectiontring = ""
    conn.Open

    For row = startRow To lasRow
        
        PopulateDeliverables "Test Plan", currentWR, row, conn, rs
        PopulateDeliverables "Test Scenarios", currentWR, row, conn, rs
        PopulateDeliverables "Test Cases", currentWR, row, conn, rs
        PopulateDeliverables "Test Results", currentWR, row, conn, rs
        PopulateDeliverables "Sign Off", currentWR, row, conn, rs
    
    Next
    
    UpdateDeliverableDates = True
    
    Exit Function

ConnectionHandler:
    
    MsgBox "Unable to connect to database with given credentials"
    UpdateDeliverableDates = False

End Function

Sub PopulateDeliverables(indicator As String, currentWR As String, row As Integer, conn As Object, rs As Object)

    Dim testPlanArray As Variant
    Dim testScenariosArray As Variant
    Dim testCasesArray As Variant
    Dim testResultsArray As Variant
    Dim signOffArray As Variant
    Dim i As Integer: i = 0
    Dim infoAvail As Boolean: infoAvail = False
    Dim itemArray As Variant
    Dim colSpot As String

    ' Creating arrays for deliverables
    If indicator = "Test Plan" Then
        testPlanArray = Array("UAT CoE Test Plan (PCR) - 3", "UAT CoE Test Plan (PCR) - 2", "UAT CoE Test Plan (PCR) - 1", _
            "UAT CoE Test Plan (PCR)", "UAT CoE Test Plan")
        itemArray = testPlanArray
        colSpot = "H"
    ElseIf indicator = "Test Cases" Then
        testCasesArray = Array("UAT CoE Test Cases FINAL", "UAT CoE Test Cases (PCR) - 3", "UAT CoE Test Cases (PCR) - 2", "UAT CoE Test Cases (PCR) - 1", _
            "UAT CoE Test Cases (PCR)", "UAT CoE Test Cases -1", "UAT CoE Test Cases")
        itemArray = testCasesArray
        colSpot = "J"
    ElseIf indicator = "Test Scenarios" Then
        testScenariosArray = Array("UAT CoE Test Scenarios (PCR) - 3", "UAT CoE Test Scenarios (PCR) - 2", "UAT CoE Test Scenarios (PCR) - 1", _
            "UAT CoE Test Scenarios (PCR)", "UAT CoE Test Scenarios -1", "UAT CoE Test Scenarios")
        itemArray = testScenariosArray
        colSpot = "I"
    ElseIf indicator = "Test Results" Then
        testResultsArray = Array("UAT CoE Test Results Updated", "UAT CoE Test Result (PCR) - 3", "UAT CoE Test Result (PCR) - 2", "UAT CoE Test Result (PCR) - 1", _
            "UAT CoE Test Result (PCR)", "UAT CoE Test Result -1", "UAT CoE Test Result", "UAT CoE test results")
        itemArray = testResultsArray
        colSpot = "K"
    ElseIf indicator = "Sign Off" Then
        signOffArray = Array("UAT COE Sign Off - Updated", "UAT CoE SignOff Updated", "UAT CoE Signoff Updated", "UAT CoE Sign Off updated", "UAT CoE Sign Off (PCR) - 3", "UAT CoE Sign Off (PCR) - 2", _
            "UAT CoE Sign Off (PCR) - 1", "UAT CoE Sign Off (PCR)", "UAT CoE Sign Off")
        itemArray = signOffArray
        colSpot = "L"
    End If

    ' Test Plan Query
    With ThisWorkbook.Worksheets("Write Tool")
        For Each Item In itemArray
            ' If complete select date_completed
            rs.Open "Select ... from ... where ... = '" & currentWR & "' and ... = '" & Item & "' with ur", conn
                If Not (rs.EOF Or rs.BOF) Then
                    If rs.Fields.Count > 0 Then
                        If Len(rs.Fields(0).Value) > 0 Then
                            .Range(colSpot & row).Value = rs.Fields(1).Value & " - Due: " & rs.Fields(0).Value
                        Else
                            .Range(colSpot & row).Value = rs.Fields(1).Value & " - No Due Date"
                        End If
                        infoAvail = True
                        rs.Close
                        Exit For
                    End If
                End If
            rs.Close
        Next
        ' Letting user know if nothing was in DB2
        If infoAvail = False Then
            .Range(colSpot & row).Value = "No info on CQ"
        End If
    End With

End Sub

Sub RemovePreviousColorsAndFilters(MasterTracker As Workbook, redValue As Integer, greenValue As Integer, blueValue As Integer)

    Dim ws As Variant
    Dim lastSheetRow As Long
    Dim row As Long
    
    For Each ws In MasterTracker.Worksheets
        With ws
            lastSheetRow = .Cells(.Rows.Count, "C").End(xlUp).row
            
            ' Remove previous intake color
            For row = 2 To lastSheetRow
                If .Range("A" & row & ":F" & row).Interior.Color = RGB(redValue, greenValue, blueValue) Then
                    .Range("A" & row & ":F" & row).Interior.ColorIndex = 0
                End If
            Next
            
            ' Remove filters
            If ws.FilterMode Then ws.ShowAllData
            If ws.AutoFilterMode Then ws.AutoFilterMode = False

        End With
    Next
    
End Sub

Function AttachMasterTracker(wsName As String, cellValue As String) As Workbook

    Dim username    As String
    Dim nameOfDoc   As String
    Dim check       As Boolean
    Dim nameLen     As Integer
    Dim wbName      As String
    
    ' Grabbing user information
    With ThisWorkbook.Worksheets(wsName)

        wbName = returnFileName(cellValue, "\")
        
        ' First checking if open
        If BookOpen(wbName) Then
            Set AttachMasterTracker = Workbooks(wbName)
        Else
            On Error GoTo FileOpener1
            Application.DisplayAlerts = False
            If InStr(cellValue, "\") Then
                Workbooks.Open Filename:=CStr(cellValue), UpdateLinks:=3
            Else
                Workbooks.Open Filename:="C:\Users\" & Environ("username") & "\Desktop\" & wbName, UpdateLinks:=3
            End If
            On Error GoTo 0
            Application.DisplayAlerts = True
            ActiveWindow.Visible = True
            ThisWorkbook.Activate
            Set AttachMasterTracker = Workbooks(wbName)
        End If
    
    End With
    
    Exit Function
    
FileOpener1:
    MsgBox "Unable to open file at path " & cellValue
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    
End Function

Function returnFileName(filePath As String, delim As String) As String

    Dim words() As String

    If InStr(filePath, delim) Then
        words = Split(filePath, delim)
        returnFileName = Trim(CStr(words(UBound(words) - LBound(words))))
    Else
        returnFileName = Trim(CStr(filePath))
    End If
    
End Function
