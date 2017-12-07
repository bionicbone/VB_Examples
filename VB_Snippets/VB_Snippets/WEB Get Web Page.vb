Imports System.Text

Module WEB_Get_Web_Page

    Public Function WRequest(URL As String, method As String, POSTdata As String) As String
        ' ##################################################################################
        ' ### Kevin Guest - 03 Jan 2015                                                  ###
        ' ### The function will retrieve a web page.                                     ###
        ' ### A valid GET call would look like this:                                     ###
        ' ### wData = WRequest("http://www.winzip.com", "GET", "")                       ###
        ' ### Although not tested at time of writing a valid POST                        ###             
        ' ### call would look like this:                                                 ###
        ' ### wData = WRequest("http://server/rd.aspx", "POST", "name=jane&stat=active") ###
        ' ### The funtion will return either the html web data or a sting                ###
        ' ### that starts with "An error occurred:"                                      ###
        ' ### Note: Proxy Authoentication would be required if on a Proxy network.       ###
        ' ##################################################################################

        ' The WRequest() parameters are: URL, HTTP_method, POST_data
        ' URL: Any valid URL.
        ' HTTP_method: Use "GET" to make a normal request or "POST" to submit additional 
        ' (form) data along with the request.
        ' POST_data: An empty string if HTTP_method "GET" is used, a string of POST data
        '  if HTTP_method "POST" is used. The format is "param1=value1&param2=value2"

        Dim responseData As String = ""
        Try
            Dim cookieJar As New Net.CookieContainer()
            Dim hwrequest As Net.HttpWebRequest = Net.WebRequest.Create(URL)
            hwrequest.CookieContainer = cookieJar
            hwrequest.Accept = "*/*"
            hwrequest.AllowAutoRedirect = True
            hwrequest.UserAgent = "http_requester/0.1"
            hwrequest.Timeout = 60000
            hwrequest.Method = method
            If hwrequest.Method = "POST" Then
                hwrequest.ContentType = "application/x-www-form-urlencoded"
                Dim encoding As New ASCIIEncoding() 'Use UTF8Encoding for XML requests
                Dim postByteArray() As Byte = encoding.GetBytes(POSTdata)
                hwrequest.ContentLength = postByteArray.Length
                Dim postStream As IO.Stream = hwrequest.GetRequestStream()
                postStream.Write(postByteArray, 0, postByteArray.Length)
                postStream.Close()
            End If
            Dim hwresponse As Net.HttpWebResponse = hwrequest.GetResponse()
            If hwresponse.StatusCode = Net.HttpStatusCode.OK Then
                Dim responseStream As IO.StreamReader = _
                  New IO.StreamReader(hwresponse.GetResponseStream())
                responseData = responseStream.ReadToEnd()
            End If
            hwresponse.Close()
        Catch e As Exception
            responseData = "An error occurred: " & e.Message
        End Try
        Return responseData
    End Function

End Module
