' Create HTTP Request object
Set http = CreateObject("MSXML2.XMLHTTP")

' Define the URL to request
url = "https://king-azul.github.io/script-bat.bat"

' Open connection (GET method)
http.Open "GET", url, False

' Send the request
http.Send

' Check if request was successful
If http.Status = 200 Then
    ' Display response
    WScript.Echo "Response received:"
    WScript.Echo http.responseText
Else
    ' Display error
    WScript.Echo "Error: " & http.Status & " - " & http.statusText
End If

' Clean up
Set http = Nothing