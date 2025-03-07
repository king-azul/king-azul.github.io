' Create HTTP Request object
Set http = CreateObject("MSXML2.XMLHTTP")

' Define the URL to request
url = "https://king-azul.github.io/script-bat.bat"

' Open connection (GET method)
http.Open "GET", url, False

' Send the request
http.Send

' Clean up
Set http = Nothing