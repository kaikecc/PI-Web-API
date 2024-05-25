' This function encodes a string in Base64 format.
' The function takes one argument: the string to encode.
' The function returns a string containing the Base64-encoded string.
Function Base64Encode(ByVal sText As String) As String
    On Error GoTo ErrorHandler
    
    Dim oXML As Object
    Dim oNode As Object
    
    ' Create an instance of the XML DOMDocument object
    Set oXML = CreateObject("Msxml2.DOMDocument")
    
    ' Create a new element for the Base64-encoded string
    Set oNode = oXML.createElement("base64")
    
    ' Set the data type of the element to binary Base64
    oNode.DataType = "bin.base64"
    
    ' Set the node value to the binary representation of the input string
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    
    ' Get the text value of the Base64-encoded string element
    Base64Encode = oNode.Text
    
    ' Clean up the objects
    Set oNode = Nothing
    Set oXML = Nothing
    
    Exit Function

ErrorHandler:
    ' If an error occurs, return an error message
    Base64Encode = "Error encoding in Base64: " & Err.Description
End Function

' This function converts a string to a binary stream.
' The function takes one argument: the string to convert.
' The function returns a variant containing the binary stream.
Function Stream_StringToBinary(ByVal sText As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim ado As Object
    
    ' Create an instance of the ADO Stream object
    Set ado = CreateObject("ADODB.Stream")
    
    ' Set the stream type to binary
    ado.Type = 2
    
    ' Set the character set to US-ASCII
    ado.Charset = "us-ascii"
    
    ' Open the stream
    ado.Open
    
    ' Write the input string to the stream
    ado.WriteText sText
    
    ' Set the stream position to the beginning
    ado.Position = 0
    
    ' Set the stream type to binary
    ado.Type = 1
    
    ' Read the binary stream into a variant
    Stream_StringToBinary = ado.Read
    
    ' Clean up the object
    Set ado = Nothing
    
    Exit Function

ErrorHandler:
    ' If an error occurs, return a null value
    Stream_StringToBinary = Null
End Function

' This function sends an HTTP GET request to an API endpoint and returns the response.
' The function takes one argument: the URL of the API endpoint.
' The function returns an array containing the API response and status code.
Function GetAPIResponse(ByVal URL As String) As String()
    On Error GoTo ErrorHandler
    
    Dim xmlHttp As Object
    Dim response(1) As String
    
    Dim username As String
    Dim password As String
    
    ' Get the username and password from the credentials form
    username = credentials.text_user.value
    password = credentials.text_pass.value
    
    ' Create an instance of the XMLHTTP object
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Open a connection to the API endpoint with Basic Authentication
    xmlHttp.Open "GET", URL, False
    xmlHttp.SetRequestHeader "Authorization", "Basic " & Base64Encode(username & ":" & password)
    
    ' Send the HTTP request and retrieve the response
    xmlHttp.Send
    
    ' Store the response text and status code in the response array
    response(0) = xmlHttp.ResponseText
    response(1) = CStr(xmlHttp.Status)
    
    ' Return the API response
    GetAPIResponse = response
    
    Exit Function

ErrorHandler:
    ' If an error occurs, return an error message and status code 404
    response(0) = "Request error: " & Err.Description
    response(1) = "404"
    GetAPIResponse = response
End Function
