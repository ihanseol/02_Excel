Attribute VB_Name = "Reference_WiseOwl_Selenium"
Option Explicit

Private cd As Selenium.ChromeDriver

Sub FindingElements()

    Set cd = New Selenium.ChromeDriver
    
    cd.Start
    cd.Get "https://en.wikipedia.org/wiki/Main_Page"
    
'=================================================================
'Find Element By ID or Name
'=================================================================
'    Dim SearchInput As Selenium.WebElement
'    Dim SearchButton As Selenium.WebElement
'    Dim FindBy As New Selenium.By
'
'    If Not cd.IsElementPresent(FindBy.ID("searchInput")) Then
'        MsgBox "Could not find search input box"
'        Exit Sub
'    End If
'
'    'Set SearchInput = cd.FindElementById("searchInput")
'    'Set SearchInput = cd.FindElement(FindBy.ID("searchInput"))
'    'Set SearchInput = cd.FindElementByName("search")
'    'Set SearchInput = cd.FindElementByCss("#searchInput")
'    'Set SearchInput = cd.FindElementByCss("[name='search']")
'    Set SearchInput = cd.FindElementByXPath("//*[@id='searchInput']")
'
'    SearchInput.SendKeys "selenium"
'
'    If Not cd.IsElementPresent(FindBy.ID("searchButton")) Then
'        MsgBox "Could not find search button"
'        Exit Sub
'    End If
'
'    'Set SearchButton = cd.FindElementById("searchButton")
'    'Set SearchButton = cd.FindElement(FindBy.ID("searchButton"))
'    'Set SearchButton = cd.FindElement(FindBy.Name("go"))
'    'Set SearchButton = cd.FindElement(FindBy.Css("#searchButton"))
'    'Set SearchButton = cd.FindElement(FindBy.Css("[name='go']"))
'    Set SearchButton = cd.FindElement(FindBy.XPath("//*[@name='go']"))
'
'    SearchButton.Click
    
'=================================================================
'Find Elements by Tag
'=================================================================
'    Dim H2Headers As Selenium.WebElements
'    Dim H2Header As Selenium.WebElement
'
'    'Set H2Headers = cd.FindElementsByTag("h2")
'    'Set H2Headers = cd.FindElementsByCss("h1, h2, h3")
'    Set H2Headers = cd.FindElementsByXPath("//h1 | //h2 | //h3")
'
'    If H2Headers.Count = 0 Then
'        MsgBox "No H2 headers found"
'        Exit Sub
'    End If
'
'    For Each H2Header In H2Headers
'        Debug.Print H2Header.tagname, H2Header.Text
'    Next H2Header
    
    
'=================================================================
'Narrowing the scope
'=================================================================
'    Dim OtdListItems As Selenium.WebElements
'    Dim OtdListItem As Selenium.WebElement
''    Dim OtdDiv As Selenium.WebElement
''    Dim OtdLists As Selenium.WebElements
''    Dim OtdList1 As Selenium.WebElement
'
''    Set OtdDiv = cd.FindElementByCss("#mp-otd")
''    Set OtdLists = OtdDiv.FindElementsByCss("ul")
''    Set OtdList1 = OtdLists(1)
''    Set OtdListItems = OtdList1.FindElementsByCss("li")
'
'    'Set OtdListItems = cd.FindElementById("mp-otd").FindElementsByTag("ul")(1).FindElementsByTag("li")
'
'    'Set OtdListItems = cd.FindElementsByCss("#mp-otd > ul:nth-of-type(1) > li")
'
'    Set OtdListItems = cd.FindElementsByXPath("//*[@id='mp-otd']/ul[1]/li")
'
'    Debug.Print OtdListItems.Count
'
'    For Each OtdListItem In OtdListItems
'        Debug.Print OtdListItem.Text
'    Next OtdListItem
    
    
'=================================================================
'Find Elements by Class
'=================================================================
'    Dim Headlines As Selenium.WebElements
'    Dim Headline As Selenium.WebElement
'
'    'Set Headlines = cd.FindElementsByCss(".MainPageBG.mp-bordered")
'    Set Headlines = cd.FindElementsByXPath("//*[@class='MainPageBG mp-bordered']")
'
'    For Each Headline In Headlines
'        Debug.Print Headline.Text
'    Next Headline
    
    
'=================================================================
'Find Elements by Link Text
'=================================================================
'    Dim OtdLinks As Selenium.WebElements
'    Dim OtdLink As Selenium.WebElement
'
'    'Set OtdLinks = cd.FindElementsByLinkText(Format(Date, "MMMM dd"))
'    'Set OtdLinks = cd.FindElementsByPartialLinkText(Format(Date, "MMMM"))
'    'Set OtdLinks = cd.FindElementsByCss("a[title^='" & Format(Date, "MMMM") & "']")
'
'    Set OtdLinks = cd.FindElementsByXPath("//a[contains(text(), 'January 19')]")
'
'    For Each OtdLink In OtdLinks
'        Debug.Print OtdLink.Attribute("href")
'    Next OtdLink
    
    'OtdLinks(1).Click
    
End Sub
