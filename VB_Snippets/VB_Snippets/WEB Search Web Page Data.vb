Module WEB_Search_Web_Page_Data

    Public Function Search_Web_Data(WebData As String, SearchUniqueBase As String, SearchUniqueBaseOffSet As Integer, SearchStart As String, SearchStartOffSet As Integer, SearchEnd As String, SearchEndOffSet As Integer) As String
        ' #############################################################
        ' ### Kevin Guest - 03 Jan 2015                             ###
        ' ### The function will search a complex string of data     ###
        ' ### like html for a specific part. Using the parameters   ###
        ' ### it is possible to jump over large parts of the string ###
        ' ### before specifying what to search for as the start     ###
        ' ### position.                                             ###
        ' ### It has been designed to pull information from a web   ###  
        ' ### page such as stock levels, prices or share details    ###
        ' #############################################################

        ' Input: WebData, SearchUniqueBase, SearchStart, SearchStartOffSet, SearchEnd, SearchEndOffSet
        '       WebData             = A string holding the response from the webpage and holds the data we need to find.
        '       SearchUniqueBase    = A unique set of characters that preceeds the start of the date we need.
        '                             Useful for jumping over a huge part of the WebPage that contains many instances of the SearchStart String.
        '       SearchStart         = A string of characters that immediately preceeds the data you need.
        '       SearchStartOffSet   = Integer Value that is the number of Charaters from the first Character of SearchStart to the String we are searching for.
        '       SearchEnd           = A string of characters that immediately Succeeds the data you need.
        '       EndStartOffSet      = Integer Value that is the number of Charaters from the first Character of SearchEnd to the String we are searching for, (usually Negative).
        '
        ' Return: The required String from the WebPage or "Failed" if the value could not be determined

        ' This is a WebPage, so I must find the actual value.

        Try
            Dim BaseCharOfSearch As Integer = 0
            Dim FirstCharOfSearch As Integer = 0
            Dim LastCharOfSearch As Integer = 0
            Dim SearchResult As String = ""


            BaseCharOfSearch = InStr(1, WebData, SearchUniqueBase) + SearchUniqueBaseOffSet
            If BaseCharOfSearch - SearchUniqueBaseOffSet <= 0 Then
                'Debug.Print("Base of Search Not Found!")
                SearchResult = "N/A"
                Return SearchResult
            End If

            FirstCharOfSearch = InStr(BaseCharOfSearch, WebData, SearchStart) + SearchStartOffSet
            If FirstCharOfSearch - SearchStartOffSet <= 0 Then
                'Debug.Print("Start of Search Not Found!")
                SearchResult = "N/A"
                Return SearchResult
            End If

            LastCharOfSearch = InStr(FirstCharOfSearch, WebData, SearchEnd) + SearchEndOffSet
            If LastCharOfSearch - SearchEndOffSet <= 0 Then
                'Debug.Print("End of Search Not Found!")
                SearchResult = "N/A"
                Return SearchResult
            End If

            SearchResult = Mid(WebData, FirstCharOfSearch, LastCharOfSearch - FirstCharOfSearch)
            'Debug.Print("Search Found: " & SearchResult)

            Return SearchResult

        Catch e As Exception
            Return "N/A"
        End Try

    End Function

End Module
