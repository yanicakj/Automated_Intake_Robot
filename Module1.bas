Attribute VB_Name = "Module1"
Option Explicit

Sub Collection()

    Dim conn As Object 
    Dim rs As Object
    Dim rs2 As Object
    Dim username As String
    Dim password As String
    Dim selectedDate As String
    Dim row As Integer: row = 8
    Dim counter As Integer
    Dim lastRow As Long
    Dim sizeByRow As Integer
    Dim workRequest As String
    Dim stateExceptions As String
    Dim initials As String
    Dim secondSheetSpot As Integer
    Dim lastRowSystems As Integer
    Dim lastRowAccImps As Integer
    Dim lastLogRow As Integer
    Dim firstsheetspot As Integer
    Dim yellowScopedItems As Collection: Set yellowScopedItems = New Collection
    Dim accImpItems As Collection: Set accImpItems = New Collection
    Dim systemsMissCount As Integer: systemsMissCount = 0
    Dim msgValue
    Dim wsName As String: wsName = "Main"
    Dim ws2Name As String: ws2Name = "Tables"
    Dim i As Integer

    ' Checking to make sure all needed information is filled out
    With ThisWorkbook.Worksheets(wsName)
        
        ' Returning username & password by ref
        If Not GetUsernameAndPassword(username, password, wsName) Then Exit Sub
        
        ' month dropdown
        If .ComboBox1.Value = "Please Select a Month" Then
            MsgBox "Please select a month to start query from!"
            .Range("H1").Activate
            Exit Sub
        End If
        
        ' day dropdown
        If .ComboBox2.Value = "Please Select a Day" Then
            MsgBox "Please select a day to start query from!"
            .Range("H3").Activate
            Exit Sub
        End If
        
        ' year dropdown
        If .ComboBox3.Value = "Please Select a Year" Then
            MsgBox "Please select a year to start query from!"
            .Range("H5").Activate
            Exit Sub
        End If
        
        Application.StatusBar = "Running..."
        
        ' Clearing previous content
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        If lastRow <> 7 Then
            .Range("A8:P" & lastRow).ClearContents
            .Range("A8:P" & lastRow).Interior.ColorIndex = 0
        End If
    End With
    
    ' Gathering intials from LAN username
    If LCase(username) = "jyanicak" Then
        initials = "JTY"
    ElseIf LCase(username) = "..." Then
        initials = ".."
    ElseIf LCase(username) = "..." Then
        initials = ".."
    ElseIf LCase(username) = "..." Then
        initials = ".."
    ElseIf LCase(username) = "..." Then
        initials = ".."
    ElseIf LCase(username) = "..." Then
        initials = ".."
    ElseIf LCase(username) = "..." Then
        initials = ".."
    Else
        initials = Mid(UCase(username), 1, 2)
    End If
    
    ' Calling sub to check for no gap in query dates
    If Not UnderLapChecker(wsName) Then
        msgValue = MsgBox("Caution. WR's may be missed with current query. Do you wish to continue?", vbYesNo)
        If msgValue = vbNo Then
            Application.StatusBar = ""
            Exit Sub
        End If
    End If
    
    'Note : timestamp('2018-05-01 00:00:00')
    With ThisWorkbook.Worksheets(wsName)
        selectedDate = .ComboBox3.Value & "-" & .ComboBox1.Value & "-" & .ComboBox2.Value
    End With

    ' Initializing connection objects
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    Set rs2 = CreateObject("ADODB.Recordset")
    
    ' Connecting to database
    On Error GoTo ConnectionHandler
    conn.ConnectionString = "Provider=IBMDADB2.1;UID=" & username & ";PWD=" & password & ";Data Source=...;ProviderType=OLEDB"
    conn.Open
    
    On Error GoTo 0
    ' Setting state exceptions string
    stateExceptions = "( state <> 16778259 or state <> 16778260 or state <> 16778261 or state <> 16778262 or state <> 16790279 or state <> 16790318 "
    stateExceptions = stateExceptions & "or state <> 16798300 or state <> 16798301 or state <> 16778327 or state <> 16788380 or state <> 16790179 "
    stateExceptions = stateExceptions & "or state <> 16783087 or state <> 16790819 or state <> 16801994 or state <> 16802006 or state <> 16806382 "
    stateExceptions = stateExceptions & "or state <> 16806803 or state <> 16803976 or state <> 16803975 or state <> 16805250 or state <> 16815640 "
    stateExceptions = stateExceptions & "or state <> 16816833 or state <> 16817902 or state <> 16823768 or state <> 16823776 or state <> 16826510 or state <> 16826518 or state <> 16833579 )"
    
    ' Queries
    With ThisWorkbook.Worksheets(wsName)
        ' First querying WR ID's
        rs.Open "Select ... " & _
                "..., " & _
                "..." & _
                "from ..." & _
                "where ... >= timestamp('" & selectedDate & " 00:00:00') " & _
                "and " & stateExceptions & " order by ... with ur", conn
        If Not (rs.BOF Or rs.EOF) Then
            Do While Not rs.EOF
                On Error GoTo Skipper2
                .Range("A" & row).Value = rs.Fields(0).Value
                .Range("B" & row).Value = rs.Fields(1).Value
                .Range("C" & row).Value = rs.Fields(2).Value
                .Range("C" & row).NumberFormat = "m/d/yyyy"
                .Range("D" & row).Value = Date
                .Range("D" & row).NumberFormat = "m/d/yyyy"
                .Range("E" & row).Value = initials
                .Range("F" & row).Value = rs.Fields(3).Value
                .Range("G" & row).Value = rs.Fields(4).Value
                .Range("H" & row).Value = rs.Fields(5).Value
                .Range("I" & row).Value = rs.Fields(6).Value
                .Range("I" & row).NumberFormat = "m/d/yyyy"
                .Range("J" & row).Value = rs.Fields(7).Value
                .Range("K" & row).Value = rs.Fields(8).Value
                On Error GoTo 0
                rs.MoveNext
                row = row + 1
            Loop
        Else
            .Range("A" & row).Value = "Empty Query"
            row = row + 1
        End If
        rs.Close
    
        ' Getting values after initial query
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        sizeByRow = lastRow - 7

    End With
    
    'Resetting screen
    ThisWorkbook.Worksheets(wsName).Range("A8").Activate
    
    'Closing connection
    conn.Close

    ' Checking for highlight-needed exceptions
    Application.StatusBar = "Checking for inscope systems"
    
    With ThisWorkbook.Worksheets(ws2Name)
        lastRowSystems = .Cells(.Rows.Count, "A").End(xlUp).row
        lastRowAccImps = .Cells(.Rows.Count, "B").End(xlUp).row
    End With
    
    With ThisWorkbook.Worksheets(wsName)
    
        ' checking system scope
        For firstsheetspot = 8 To lastRow
            For row = 1 To lastRowSystems
                If InStr(Trim(CStr(.Range("G" & firstsheetspot).Value)), Trim(CStr(ThisWorkbook.Worksheets(ws2Name).Range("A" & row).Value))) > 0 Then
                    .Range("L" & firstsheetspot).Value = "Inscope"
                    If Trim(.Range("B" & firstsheetspot).Value) = "No" Then
                        .Range("B" & firstsheetspot).Interior.Color = RGB(255, 255, 0)
                        yellowScopedItems.Add firstsheetspot
                    End If
                End If
            Next
            
            ' checking acc imps
            For row = 1 To lastRowAccImps
                If InStr(LCase(Trim(CStr(.Range("F" & firstsheetspot).Value))), LCase(Trim(CStr(ThisWorkbook.Worksheets(ws2Name).Range("B" & row).Value)))) > 0 Then
                    .Range("F" & firstsheetspot).Interior.Color = RGB(240, 155, 68)
                    .Range("F" & firstsheetspot).WrapText = True
                    .Range("J" & firstsheetspot).Value = .Range("J" & firstsheetspot).Value & vbCrLf & "Account Implementation, consult Acc Imp Leads"
                    If .Range("J" & firstsheetspot).Value <> "Withdrawn" And .Range("J" & firstsheetspot).Value <> "Closed" Then
                        accImpItems.Add firstsheetspot
                    End If
                    Exit For
                End If
            Next
        Next
        
    End With
    
    ' Call front end runner function here
    If systemsMissCount > 0 Then
        If Not FrontEndRunner(sizeByRow, wsName, password, systemsMissCount) Then Exit Sub
    End If
    
    ' Generate email to Mani & Dan with yellowScopedItems here
    If yellowScopedItems.Count > 0 Then
        EmailGenerator yellowScopedItems, "yellowSys", username, password, wsName
    End If
    
    ' Generate email to Acc Imp team with Acc Imp WRs here
    If accImpItems.Count > 0 Then
        EmailGenerator accImpItems, "accImp", username, password, wsName
    End If
    
    ' Updating Log
    With ThisWorkbook.Worksheets("Log")
        If ThisWorkbook.Worksheets(wsName).ComboBoxLog.Value = "On" Then
            lastLogRow = .Cells(.Rows.Count, "A").End(xlUp).row + 1
            .Range("A" & lastLogRow).NumberFormat = "m/d/yyyy"
            .Range("A" & lastLogRow).Value = Date
            .Range("B" & lastLogRow).NumberFormat = "m/d/yyyy"
            .Range("B" & lastLogRow).Value = selectedDate
            .Range("C" & lastLogRow).NumberFormat = "@"
            .Range("C" & lastLogRow).Value = initials
        End If
    End With
    
    ' Cleaning up
    With ThisWorkbook.Worksheets(wsName)
        .Range("A1").Activate
        .Range("A8").Activate
    End With
    
    Application.StatusBar = ""
    MsgBox "Done!"
    Exit Sub
    
Skipper2:
    ThisWorkbook.Worksheets(wsName).Range("G" & row).Interior.Color = RGB(0, 125, 255)
    systemsMissCount = systemsMissCount + 1
    Resume Next
    
ConnectionHandler:
    ThisWorkbook.Worksheets(wsName).Range("F2").Activate
    MsgBox "Unable to connect to database. Please make sure credentials are correct!"
    
End Sub

Function FrontEndRunner(sizeByRow As Integer, wsName As String, password As String, systemsMissCount As Integer) As Boolean

    Dim i As Integer: i = 0
    Dim ie As InternetExplorer
    Dim objCollection As Object
    Dim loginObj As Object
    Dim ObjElement As Object
    Dim searchStringObj As Object
    Dim counter As Integer: counter = 0
    Dim closeObj As Object
    Dim logoutObj As Object
    Dim pwObj As Object
    Dim checkedSoFar As Integer: checkedSoFar = 0
    
    Application.StatusBar = "Front End Grabbing Systems"
    FrontEndRunner = True
    
    On Error GoTo ErrHandler2
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = False
    
    On Error Resume Next
    Application.StatusBar = "Loading ..."
    ie.navigate "http://..."
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:07"))
    Application.StatusBar = "Please wait..."
    
    ' Checking if user is logged in
    Set objCollection = ie.document.getElementsByTagName("input")
    Do While i < objCollection.Length
        If objCollection(i).Name = "passwordId" Then
            objCollection(i).Value = password
            Set loginObj = ie.document.getElementById("...")
            loginObj.Click
            Application.Wait (Now + TimeValue("0:00:05"))
            Exit Do
        End If
        i = i + 1
    Loop
    
    ' Searching for "search" bar and saving as object
    i = 0
    Do While i < objCollection.Length
        If objCollection(i).Name = "..." Then
            Set searchStringObj = objCollection(i)
            Exit Do
        End If
        i = i + 1
    Loop
    
    ' Saving search button
    Set ObjElement = ie.document.getElementById("...")

    ' Starting main loop of checking WRs
    For counter = 0 To sizeByRow
        If ThisWorkbook.Worksheets(wsName).Range("G" & (counter + 8)).Interior.Color = RGB(0, 125, 255) Then
            ' Calling function to grab elements from html
            IntakeLooper counter, sizeByRow, ie, searchStringObj, ObjElement, checkedSoFar, wsName, systemsMissCount
            checkedSoFar = checkedSoFar + 1
        End If
    Next
    
    ' Letting processes run out
    Application.StatusBar = "Finishing Process..."
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    
    ' Closing down CQ
    Set closeObj = ie.document.getElementsByClassName("...")(0)
    closeObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    
    ' Logging out
    Set logoutObj = ie.document.getElementById("...")
    logoutObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:03"))

    ' Logging back in
    Set pwObj = ie.document.getElementById("...")
    pwObj.Value = password
    Set loginObj = ie.document.getElementById("...")
    loginObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:06"))
    
    ' Cleanup
    ie.Quit
    Set ie = Nothing
    Set ObjElement = Nothing
    Set objCollection = Nothing
    Application.StatusBar = ""
    Exit Function

ErrHandler2:
    MsgBox "Please wait a couple seconds. IE is still closing from previous use. Please restart Wait, Open & Close Worksheet, or Force Shutdown Internet Explorer"
    FrontEndRunner = False

End Function

Sub IntakeLooper(counter As Integer, sizeByRow As Integer, ie As InternetExplorer, searchStringObj As Object, ObjElement As Object, checkedSoFar As Integer, wsName As String, systemsMissCount As Integer)

    Dim closeObj        As Object
    Dim systemsObj      As Object
    Dim systems         As String
    Dim systemsID       As String

    ' Updating the status bar programmatically
    Application.StatusBar = "Checking " & CStr(ThisWorkbook.Worksheets(wsName).Range("A" & CStr(counter + 8)).Value) & ", " & CStr(Round(checkedSoFar / systemsMissCount, 2) * 100) & "% Complete"
    
    ' Searching WR
    searchStringObj.Value = CStr(ThisWorkbook.Worksheets(wsName).Range("A" & CStr(counter + 8)).Value)
    ObjElement.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:06"))

    ' Finding Systems Tested
    systemsID = "..." & CStr(1 + (checkedSoFar * 4))
    On Error Resume Next
    Set systemsObj = ie.document.getElementById(systemsID)
    systems = systemsObj.innerText
    With ThisWorkbook.Worksheets(wsName).Range("G" & (counter + 8))
        .Value = systems
        .Activate
        .Interior.ColorIndex = 0
    End With
    
    ' Closing tab
    Set closeObj = ie.document.getElementsByClassName("...")(0)
    closeObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:01"))
    
End Sub

Function UnderLapChecker(wsName As String) As Boolean

    Dim tempDate As Date
    Dim logDate As Date
    Dim month As Integer
    Dim day As Integer
    Dim year As Integer
    Dim lastLogRow As Integer
    
    With ThisWorkbook.Worksheets("Log")
        lastLogRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    
    With ThisWorkbook.Worksheets(wsName)
        month = .ComboBox1.Value
        day = .ComboBox2.Value
        year = .ComboBox3.Value
    End With
    
    tempDate = DateSerial(year, month, day)
    logDate = ThisWorkbook.Worksheets("Log").Range("A" & lastLogRow).Value
    
    If tempDate > logDate Then
        UnderLapChecker = False
    Else
        UnderLapChecker = True
    End If

End Function

Sub EmailGenerator(tempCollection As Collection, indicator As String, username As String, password As String, wsName As String)

    Dim olApp As Outlook.Application
    Dim newMail As Outlook.MailItem
    Dim emailBody As String
    Dim emailTable As String
    Dim emailSig As String
    Dim workRequest As String
    Dim uatInv As String
    Dim submitDate As String
    Dim pulledDate As String
    Dim initials As String
    Dim systemList As String
    Dim itLead As String
    Dim releaseDate As String
    Dim headline As String
    Dim reqDiv As String
    Dim lastRow As Integer
    Dim addedHeaders As Boolean: addedHeaders = False
    Dim conn As Object 'Variable for ADODB.Connection object
    Dim rs As Object 'Variable for ADODB.Recordset object
    Dim hexColor As String
    Dim tempState As String
    Dim tempDiv As String
    Dim row As Variant
    Dim tdHTML As String: tdHTML = "<td style=""border: 1px solid black;border-collapse:collapse;padding:10px"">"
    Dim secondSheetSpot As Integer
    Dim ws2Name As String: ws2Name = "Tables"
    
    ' Instantiating connection objects
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo ConnectionHandler2
    conn.ConnectionString = "Provider=IBMDADB2.1;UID=" & username & ";PWD=" & password & ";Data Source=...;ProviderType=OLEDB"
    conn.Open
    
    On Error GoTo 0
    ' Creating new email and finding last row
    Set olApp = CreateObject("Outlook.Application")  'Another option: Set olApp = New Outlook.Application
    Set newMail = olApp.CreateItem(olMailItem)
    
    With ThisWorkbook.Worksheets(wsName)
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    
    ' Starting email body html string
    If indicator = "yellowSys" Then
        emailBody = "<p style=""font-family:calibri;font-size:11pt"">Hi Guys,</p><p style=""font-family:calibri;font-size:11pt"">The below Work Requests have <b>UAT Inv = ""No""</b> but Systems to be Tested are <b>in-scope</b>.</p>"
        hexColor = "eafbfb"
    ElseIf indicator = "accImp" Then
        emailBody = "<p style=""font-family:calibri;font-size:11pt"">Hi Team,</p><p style=""font-family:calibri;font-size:11pt"">Just an FYI of the possible Account Implementation WRs that came in through the intake this past week:</p>"
        hexColor = "ffb3b3"
    ElseIf indicator = "notTesting" Then
        emailBody = "<p style=""font-family:calibri;font-size:11pt"">Hi Guys,</p><p style=""font-family:calibri;font-size:11pt"">These WRs came up as <b>UAT Inv = ""Yes""</b> under the <b>Not Testing</b> tab on Master Tracker:</p>"
        hexColor = "ccffcc"
    ElseIf indicator = "tbdTab" Then
        emailBody = "<p style=""font-family:calibri;font-size:11pt"">Hi Guys,</p><p style=""font-family:calibri;font-size:11pt"">Included are flagged WRs from the <b>TBD tab</b> of the Master Tracker:</p>"
        hexColor = "fbff6d"
    End If
    
    ' Iterating through rows in collection
    For Each row In tempCollection
    
        If indicator = "yellowSys" Or indicator = "accImp" Then
            
            ' Grabbing info from main tab
            With ThisWorkbook.Worksheets(wsName)
                workRequest = .Range("A" & row).Value
                uatInv = .Range("B" & row).Value
                submitDate = .Range("C" & row).Value
                pulledDate = .Range("D" & row).Value
                initials = .Range("E" & row).Value
                headline = .Range("F" & row).Value
                systemList = .Range("G" & row).Value
                itLead = .Range("H" & row).Value
                releaseDate = .Range("I" & row).Value
            End With
            
            ' for linebreaks
            systemList = Replace(systemList, vbCrLf, "<br>")
            systemList = Replace(systemList, Chr(10), "<br>")
            systemList = Replace(systemList, Chr(13), "<br>")
            
            ' querying for requesting division from CQ
            rs.Open "Select requesting_division from cquadmp1.it_work_request where id = '" & workRequest & "' with ur", conn
            If Not (rs.BOF Or rs.EOF) Then
                reqDiv = rs.Fields(0).Value
            Else
                reqDiv = "blank"
            End If
            rs.Close
            
        ElseIf indicator = "notTesting" Then
            workRequest = row ' row is really WR here
            rs.Open "Select ... " & _
                    "from ... as t1 inner join... as t2 on ... = t2.id where t1.id = '" & workRequest & "' with ur", conn
            If Not (rs.BOF Or rs.EOF) Then
                releaseDate = rs.Fields(0).Value
                tempState = rs.Fields(1).Value
                headline = rs.Fields(2).Value
                tempDiv = rs.Fields(3).Value
            End If
            rs.Close
            
        ElseIf indicator = "tbdTab" Then
            workRequest = row ' row is really WR here
            rs.Open "Select .... " & _
                    "from ... as t1 inner join ... as t2 on t1.state = t2.... where t1.... = '" & workRequest & "' with ur", conn
            If Not (rs.BOF Or rs.EOF) Then
                releaseDate = rs.Fields(0).Value
                tempState = rs.Fields(1).Value
                headline = rs.Fields(2).Value
            End If
            rs.Close
        End If
        
        ' Yellow No format
        If indicator = "yellowSys" Then
            ' Adding headers
            If addedHeaders = False Then
                emailTable = "<table style=""border: 1px solid black;border-collapse:collapse; font-family:calibri; font-size:11pt"">"
                emailTable = emailTable & "<tbody style=""border: 1px solid black;border-collapse:collapse"">"
                emailTable = emailTable & "<tr bgcolor=""" & hexColor & """ style=""border: 1px solid black;border-collapse:collapse;padding:10px;text-align:center"">"
                emailTable = emailTable & "<td style=""border: 1px solid black;border-collapse:collapse;padding:10px""><b>WR</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>UAT Inv</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Date Submitted</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Date Pulled</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Reviewed By</b></td>"
                emailTable = emailTable & "<td  width=""25%"" style=""border: 1px solid black;padding:10px""><b>Headline</b></td>"
                emailTable = emailTable & "<td  width=""25%"" style=""border: 1px solid black;padding:10px""><b>Systems Tested</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>IT Lead</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Requesting Div.</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Release Date</b></td></tr>"
                ' New row for data
                emailTable = emailTable & "<tr style=""border: 1px solid black;border-collapse:collapse;padding:10px"">"
                emailTable = emailTable & tdHTML & workRequest & "</td>"
                emailTable = emailTable & tdHTML & uatInv & "</td>"
                emailTable = emailTable & tdHTML & submitDate & "</td>"
                emailTable = emailTable & tdHTML & pulledDate & "</td>"
                emailTable = emailTable & tdHTML & initials & "</td>"
                emailTable = emailTable & tdHTML & headline & "</td>"
                emailTable = emailTable & tdHTML & systemList & "</td>"
                emailTable = emailTable & tdHTML & itLead & "</td>"
                emailTable = emailTable & tdHTML & reqDiv & "</td>"
                emailTable = emailTable & tdHTML & releaseDate & "</td></tr>"
                addedHeaders = True
            ' add just bottom row (data)
            Else
                emailTable = emailTable & "<tr style=""border: 1px solid black;border-collapse:collapse;padding:10px"">"
                emailTable = emailTable & tdHTML & workRequest & "</td>"
                emailTable = emailTable & tdHTML & uatInv & "</td>"
                emailTable = emailTable & tdHTML & submitDate & "</td>"
                emailTable = emailTable & tdHTML & pulledDate & "</td>"
                emailTable = emailTable & tdHTML & initials & "</td>"
                emailTable = emailTable & tdHTML & headline & "</td>"
                emailTable = emailTable & tdHTML & systemList & "</td>"
                emailTable = emailTable & tdHTML & itLead & "</td>"
                emailTable = emailTable & tdHTML & reqDiv & "</td>"
                emailTable = emailTable & tdHTML & releaseDate & "</td></tr>"
            End If
            
        ' Account Imp format
        ElseIf indicator = "accImp" Then
            ' Adding headers
            If addedHeaders = False Then
                emailTable = "<table style=""border: 1px solid black;border-collapse:collapse; font-family:calibri; font-size:11pt"">"
                emailTable = emailTable & "<tbody style=""border: 1px solid black;border-collapse:collapse"">"
                emailTable = emailTable & "<tr bgcolor=""" & hexColor & """ style=""border: 1px solid black;border-collapse:collapse;padding:10px;text-align:center"">"
                emailTable = emailTable & "<td style=""border: 1px solid black;border-collapse:collapse;padding:10px""><b>WR</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Date Submitted</b></td>"
                emailTable = emailTable & "<td  width=""25%"" style=""border: 1px solid black;padding:10px""><b>Headline</b></td>"
                emailTable = emailTable & "<td  width=""25%"" style=""border: 1px solid black;padding:10px""><b>Systems Tested</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>IT Lead</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Release Date</b></td></tr>"
                ' New row for data
                emailTable = emailTable & "<tr style=""border: 1px solid black;border-collapse:collapse;padding:10px"">"
                emailTable = emailTable & tdHTML & workRequest & "</td>"
                emailTable = emailTable & tdHTML & submitDate & "</td>"
                emailTable = emailTable & tdHTML & headline & "</td>"
                emailTable = emailTable & tdHTML & systemList & "</td>"
                emailTable = emailTable & tdHTML & itLead & "</td>"
                emailTable = emailTable & tdHTML & releaseDate & "</td></tr>"
                addedHeaders = True
            ' add just bottom row (data)
            Else
                emailTable = emailTable & "<tr style=""border: 1px solid black;border-collapse:collapse;padding:10px"">"
                emailTable = emailTable & tdHTML & workRequest & "</td>"
                emailTable = emailTable & tdHTML & submitDate & "</td>"
                emailTable = emailTable & tdHTML & headline & "</td>"
                emailTable = emailTable & tdHTML & systemList & "</td>"
                emailTable = emailTable & tdHTML & itLead & "</td>"
                emailTable = emailTable & tdHTML & releaseDate & "</td></tr>"
            End If
            
        ' Not Testing - UAT Yes format
        ElseIf indicator = "notTesting" Then
            ' Adding headers
            If addedHeaders = False Then
                emailTable = "<table style=""border: 1px solid black;border-collapse:collapse; font-family:calibri; font-size:11pt"">"
                emailTable = emailTable & "<tbody style=""border: 1px solid black;border-collapse:collapse"">"
                emailTable = emailTable & "<tr bgcolor=""" & hexColor & """ style=""border: 1px solid black;border-collapse:collapse;padding:10px;text-align:center"">"
                emailTable = emailTable & "<td style=""border: 1px solid black;border-collapse:collapse;padding:10px""><b>WR</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;border-collapse:collapse;padding:10px""><b>UAT Inv</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;border-collapse:collapse;padding:10px""><b>State</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Release Date</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Project</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Requesting Div</b></td></tr>"
                ' New row for data
                emailTable = emailTable & "<tr style=""border: 1px solid black;border-collapse:collapse;padding:10px"">"
                emailTable = emailTable & tdHTML & workRequest & "</td>"
                emailTable = emailTable & tdHTML & "Yes" & "</td>"
                emailTable = emailTable & tdHTML & tempState & "</td>"
                emailTable = emailTable & tdHTML & releaseDate & "</td>"
                emailTable = emailTable & tdHTML & headline & "</td>"
                emailTable = emailTable & tdHTML & tempDiv & "</td></tr>"
                addedHeaders = True
            ' add just bottom row (data)
            Else
                emailTable = emailTable & "<tr style=""border: 1px solid black;border-collapse:collapse;padding:10px"">"
                emailTable = emailTable & tdHTML & workRequest & "</td>"
                emailTable = emailTable & tdHTML & "Yes" & "</td>"
                emailTable = emailTable & tdHTML & tempState & "</td>"
                emailTable = emailTable & tdHTML & releaseDate & "</td>"
                emailTable = emailTable & tdHTML & headline & "</td>"
                emailTable = emailTable & tdHTML & tempDiv & "</td></tr>"
            End If
            
        ' TBD Tab - Flagged format
        ElseIf indicator = "tbdTab" Then
            ' Adding headers
            If addedHeaders = False Then
                emailTable = "<table style=""border: 1px solid black;border-collapse:collapse; font-family:calibri; font-size:11pt"">"
                emailTable = emailTable & "<tbody style=""border: 1px solid black;border-collapse:collapse"">"
                emailTable = emailTable & "<tr bgcolor=""" & hexColor & """ style=""border: 1px solid black;border-collapse:collapse;padding:10px;text-align:center"">"
                emailTable = emailTable & "<td style=""border: 1px solid black;border-collapse:collapse;padding:10px""><b>WR</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;border-collapse:collapse;padding:10px""><b>UAT Inv</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;border-collapse:collapse;padding:10px""><b>State</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Release Date</b></td>"
                emailTable = emailTable & "<td style=""border: 1px solid black;padding:10px""><b>Project</b></td></tr>"
                ' New row for data
                emailTable = emailTable & "<tr style=""border: 1px solid black;border-collapse:collapse;padding:10px"">"
                emailTable = emailTable & tdHTML & workRequest & "</td>"
                emailTable = emailTable & tdHTML & "Yes" & "</td>"
                emailTable = emailTable & tdHTML & tempState & "</td>"
                emailTable = emailTable & tdHTML & releaseDate & "</td>"
                emailTable = emailTable & tdHTML & headline & "</td></tr>"
                addedHeaders = True
            ' add just bottom row (data)
            Else
                emailTable = emailTable & "<tr style=""border: 1px solid black;border-collapse:collapse;padding:10px"">"
                emailTable = emailTable & tdHTML & workRequest & "</td>"
                emailTable = emailTable & tdHTML & "Yes" & "</td>"
                emailTable = emailTable & tdHTML & tempState & "</td>"
                emailTable = emailTable & tdHTML & releaseDate & "</td>"
                emailTable = emailTable & tdHTML & headline & "</td></tr>"
            End If
        
        End If
    Next
    
    ' Closing table
    emailTable = emailTable & "</tbody></table>"
    
    emailSig = "<p style=""font-family:calibri;font-size:11pt"">Thank you,<br><br>Jacob Yanicak<br>UAT-COE Analyst<br>Horizon Blue Cross Blue Shield of NJ<br>....com</p>"
    ' Putting everything back together
    emailBody = emailBody & "<br>" & emailTable & "<br>" & emailSig
        
    With newMail
        If indicator = "yellowSys" Or indicator = "notTesting" Or indicator = "tbdTab" Then
            .To = "....com; ....com"
        ElseIf indicator = "accImp" Then
            .To = "UAT-....com"
            .CC = "....com; ...com"
        End If
        If indicator = "yellowSys" Then
            .Subject = "Intake Work Requests UAT = 'No' with Systems In-scope"
        ElseIf indicator = "accImp" Then
            .Subject = "Past Week's Account Implementation WRs"
        ElseIf indicator = "notTesting" Then
            .Subject = "WRs with UAT = 'Yes' under Not Testing Tab"
        ElseIf indicator = "tbdTab" Then
            .Subject = "Flagged WRs from TBD Tab"
        End If
        .SentOnBehalfOfName = "....com"
        .HTMLBody = emailBody
        .Save
        .Close olPromptForSave
    End With
    
    conn.Close
    Exit Sub
    
ConnectionHandler2:
    ThisWorkbook.Worksheets(wsName).Range("F2").Activate
    MsgBox "Unable to connect to database"

End Sub
