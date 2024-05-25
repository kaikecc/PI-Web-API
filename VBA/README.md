
# PI Web API Project with VBA

This project is a VBA (Visual Basic for Applications) script developed to extract data from a PI Web API and write that data into an Excel spreadsheet. It uses basic authentication to communicate with the PI Web API and extract the data. The goal of this code is not to replace PI Builder or PI DataLink but to generate insights for automations using VBA macros.

## Prerequisites

To run this script, you need the following:

* Microsoft Excel (with VBA support)
* Access to a PI Web API
* Import the **JsonConverter** file
* Referências adicionadas no Editor VBA para:
  * Microsoft Scripting Runtime
  * Microsoft WinHTTP Services, version 5.1

## Usage

The script defines two main subroutines:

1. **ExtractPIWebAPI(endpoint As String):** This is the main subroutine that coordinates the extraction of data from the PI Web API and writes that data into the Excel spreadsheet. This subroutine takes one input argument:

   * endpoint: The URL of the starting point of the PI Web API hierarchy from which you intend to extract data, e.g.: https://myserver/piwebapi/webid/elements


2. **GetAPIResponse(url As String) As String:** This auxiliary function is used by the main subroutine to make requests to the PI Web API and get responses. The function also takes one input argument and returns the PI Web API response as a string. You need to set the credentials in this function.    

   * username: The username for authentication with the PI Web API
   * password: The password for authentication with the PI Web API

The script also defines two additional functions, Base64Encode(sText As String) As String and Stream_StringToBinary(sText As String) As Variant, which are used to encode the authentication string in Base64.

## How to Run the Script

To run the script:

1. Open Excel and access the VBA Editor (Shortcut: ALT + F11).
2. In the VBA Editor, import this script.
3. Add the necessary references (Microsoft Scripting Runtime and Microsoft WinHTTP Services, version 5.1) through the "Tools" -> "References" menu.
4. In your VBA code, call the ExtractPIWebAPI(endpoint, username, password) subroutine with the correct credentials and endpoint.
5. Run your VBA code.

When executed, the script extracts data from the specified PI Web API, processes that data, and writes the results into the "PI Tags" worksheet of the current Excel file.

## Considerations

Ensure that the provided username and password have the correct permissions to access the data in the PI Web API.

This code is not optimized for large volumes of data and may take time to execute on large datasets. If you are dealing with large volumes of data, you may need to optimize or modify this script for better performance.

Finally, this script was developed for use with a specific PI Web API and may not work correctly with all PI Web APIs. If you are having issues, check if the PI Web API is functioning correctly and if the data you are trying to extract is available.
