Sub GoogleAutomatedSearch()
    On Error GoTo ErrorHandler
    
    If MsgBox("Are you sure to start scraping?", vbQuestion + vbYesNo, "Google Listing Scraper") = vbNo Then Exit Sub
    
    
    Dim driver As New WebDriver
    Dim googleURL As String, searchText As String
    Dim lastRow As Integer, beginRow As Integer, endRow As Integer
    Dim companyName As String, companyWebsite As String, companyAddress As String, companyPhone As String
    Dim i As Integer
    
    
    Application.DisplayStatusBar = True
    
    ''' Get total number of rows to scrap
    lastRow = Sheets("Google").Range("A" & Rows.Count).End(xlUp).Row
    
    
    ''' Get Start Row and End Row number entered in respected cells
    beginRow = Val(Sheets("Google").Range("H4").Value)
    endRow = Val(Sheets("Google").Range("H6").Value)
        
        
    ''' Open Chrome browser with Visible=False
    driver.AddArgument "--headless"
    driver.Start "Chrome"
    
    
    '' Begin row must start with 3
    If beginRow < 3 Then beginRow = 3
    
    '''Loop through all given range of rows
    For i = beginRow To endRow
        '''Get Search text
        searchText = Sheets("Google").Range("A" & i).Value
        
        '' If any empty cell found then end script
        If searchText = "" Then Exit For
        
        With Application.WorksheetFunction
            ''' Replace Following special symbol with their URL compatible string
            ''' 1. space with %20
            ''' 2. & with %26
            ''' 3. , with %2C
            searchText = .Substitute(.Substitute(.Substitute(searchText, " ", "%20"), "&", "%26"), ",", "%2C")
        End With
        
        googleURL = "https://www.google.com/search?q=" & searchText
        
        
        ''Show status on Excel status bar
        Debug.Print googleURL
        Application.StatusBar = "Macro is running ... Now at row : " & i & " / " & endRow & "... Last search made at : " & Now

        
        driver.Get googleURL
        
        
        '''''Scrap Data
        companyName = ""
        companyWebsite = ""
        companyAddress = ""
        companyPhone = ""
        
        
        ''' Scrap Company Name
        companyName = driver.FindElementById("rhs").FindElementByClass("SPZz6b").FindElementByTag("h2").Attribute("innerText")
        
        Dim attr As Variant
        
        ''''Scrap Company Website URL
        For Each attr In driver.FindElementsByClass("QqG1Sd")
            If attr.Attribute("innerText") = "Website" Then
                companyWebsite = attr.Attribute("innerHTML")
                companyWebsite = Split(companyWebsite, "href=""")(1)
                companyWebsite = Split(companyWebsite, """")(0)
            End If
        Next attr
                
             
        Dim tmpText As String
        
        ''' Scrap Company Address and Phone Number
        For Each attr In driver.FindElementsByCss("div[class='zloOqf PZPZlf']")
            tmpText = attr.Attribute("innerText")
            
            If InStr(tmpText, "Address: ") > 0 Then
                companyAddress = tmpText
                companyAddress = Split(companyAddress, "Address: ")(1)
            End If
            
            If InStr(tmpText, "Phone: ") > 0 Then
                companyPhone = tmpText
                companyPhone = Split(companyPhone, "Phone: ")(1)
            End If
        Next attr
        
        '''' Stored scraped data in respected cell and Save workbook
        Sheets("Google").Range("B" & i).Value = companyName
        Sheets("Google").Range("C" & i).Value = companyAddress
        Sheets("Google").Range("D" & i).Value = companyPhone
        Sheets("Google").Range("E" & i).Value = companyWebsite
        ThisWorkbook.Save
        
        Application.Wait (Now() + TimeValue("00:00:" & Sheets("Google").Range("H8").Value))
    Next i

    Application.StatusBar = ""
    Sheets("Google").Activate
    
    driver.Quit
    Set driver = Nothing
    
    MsgBox "All keywords are searched and scrapped successfully!", vbInformation, "Google Listing Scraper"
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Google Listing Scraper"
End Sub
