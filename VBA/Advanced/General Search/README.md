# PI Web API Project with VBA

This project is a VBA (Visual Basic for Applications) script developed to extract data from a PI Web API and write that data into an Excel spreadsheet. It uses basic authentication to communicate with the PI Web API and extract the data. The goal of this code is not to replace PI Builder or PI DataLink but to generate insights for automations using VBA macros.

## Prerequisites

To run this script, you need the following:

* Microsoft Excel (with VBA support)
* Access to a PI Web API
* Import the JsonConverter file
* Added references in the VBA Editor for:
* Microsoft Scripting Runtime
* Microsoft WinHTTP Services, version 5.1
  
## Usage

The script defines two main subroutines:

1. AdvancedSearch(endpoint As String, path As String, element As String, databaseWebID As String, typeFilter As String, AFObject As String, wsName As String, pointSource As String): This is the main subroutine that coordinates the extraction of data from the PI Web API and writes that data into the Excel spreadsheet. This subroutine takes several input arguments:

* endpoint: The URL of the starting point of the PI Web API hierarchy from which you intend to extract data.
* path: The path in the AF database.
* element: The element to filter on.
* databaseWebID: The WebID of the AF database.
* typeFilter: The type of filter to apply.
* AFObject: The AF object type (e.g., "attributes", "elements", "analyses", "points").
* wsName: The name of the worksheet where the data will be written.
* pointSource: The point source filter.

2. ProcessElement(item As Variant, ws As Worksheet, Data() As Variant, ByRef rows As Long): This auxiliary function is used by the main subroutine to process element data and write it to the worksheet.

3. ProcessAttribute(item As Variant, ws As Worksheet, Data() As Variant, ByRef rows As Long, current_value_mode As Boolean): This auxiliary function is used by the main subroutine to process attribute data and write it to the worksheet.

4. ProcessAnalysis(item As Variant, ws As Worksheet, Data() As Variant, ByRef rows As Long): This auxiliary function is used by the main subroutine to process analysis data and write it to the worksheet.

5. ProcessPoints(item As Variant, ws As Worksheet, Data() As Variant, ByRef rows As Long): This auxiliary function is used by the main subroutine to process point data and write it to the worksheet.

## How to Run the Script

To run the script:

1. Open Excel and access the VBA Editor (Shortcut: ALT + F11).
2. In the VBA Editor, import this script.
3. Add the necessary references (Microsoft Scripting Runtime and Microsoft WinHTTP Services, version 5.1) through the "Tools" -> "References" menu.
4. In your VBA code, call the ExtractQueries subroutine with the correct parameters.
5. Run your VBA code.
6. 
When executed, the script extracts data from the specified PI Web API, processes that data, and writes the results into the specified worksheet of the current Excel file.

## Considerations

Ensure that the provided username and password have the correct permissions to access the data in the PI Web API.

This code is not optimized for large volumes of data and may take time to execute on large datasets. If you are dealing with large volumes of data, you may need to optimize or modify this script for better performance.

Finally, this script was developed for use with a specific PI Web API and may not work correctly with all PI Web APIs. If you are having issues, check if the PI Web API is functioning correctly and if the data you are trying to extract is available.

## Code Overview

The main subroutine AdvancedSearch handles the extraction of data from the PI Web API and writes it to an Excel worksheet. It disables automatic calculation and screen updating for better performance. The subroutine constructs the API query URL based on the specified parameters and makes a request to the PI Web API. The response is parsed, and the data is processed and written to the worksheet.

