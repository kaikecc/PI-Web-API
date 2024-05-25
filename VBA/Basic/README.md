# VBA Helper Functions for API Integration

This README provides an overview and usage instructions for a set of VBA helper functions designed to assist with API integration, including Base64 encoding, string-to-binary conversion, and making HTTP GET requests to an API endpoint.

## Functions Overview

1. Base64Encode
This function encodes a given string into Base64 format.

Prototype:

```vba
Function Base64Encode(ByVal sText As String) As String
```

Parameters:

* sText (String): The input string to encode.
  
Returns:

* (String): The Base64-encoded string.
  
Example:

```vba
Dim encodedString As String
encodedString = Base64Encode("Hello, World!")
```

Details:

This function uses an XML DOMDocument to convert the input string into a binary stream and then encodes it into Base64.

Error Handling:

If an error occurs during encoding, the function returns an error message.


2. Stream_StringToBinary
This function converts a given string into a binary stream.

Prototype:

```vba
Function Stream_StringToBinary(ByVal sText As String) As Variant
```

Parameters:

sText (String): The input string to convert.
Returns:

(Variant): The binary stream representation of the input string.

Example:

```vba
Dim binaryStream As Variant
binaryStream = Stream_StringToBinary("Hello, World!")
```

Details:

This function uses an ADO Stream object to write the input string and then reads it back as a binary stream.

Error Handling:

If an error occurs during the conversion, the function returns a null value.

3. GetAPIResponse
   
This function sends an HTTP GET request to an API endpoint and returns the response and status code.

Prototype:

```vba
Function GetAPIResponse(ByVal URL As String) As String()
```

Parameters:

URL (String): The URL of the API endpoint.
Returns:

(String Array): An array containing the API response and status code.

Example:


```vba
Dim response() As String
response = GetAPIResponse("https://api.example.com/piwebapi/")
```