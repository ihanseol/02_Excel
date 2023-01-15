Attribute VB_Name = "naver_movie_rank"
Sub get_naver_movie_ranking()

    Dim driver As WebDriver
    Dim ele As WebElement
    Dim nCnt As Integer
    
    Set driver = New WebDriver
    
    driver.Start "chrome"
    
    'wait 1 second
    driver.Wait (1000)
    
    driver.Get "https://movie.naver.com/movie/sdb/rank/rmovie.nhn"
    
    nCnt = 1
    For Each ele In driver.FindElementsByClass("tit3")
        Range("a" & nCnt).Value = ele.Text
        nCnt = nCnt + 1
    Next ele
    
    driver.Close
    Set driver = Nothing
    
End Sub
