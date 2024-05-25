Attribute VB_Name = "AdvancedSearch"
' Required References:
' Microsoft Scripting Runtime
' Microsoft WinHTTP Services, version 5.1
' JsonConverter: https://github.com/VBA-tools/VBA-JSON
' Version 7
' Autor: Kaike Castro Carvalho
' Date: 25-May-2024

Sub AdvancedSearch(ByVal endpoint As String, ByVal path As String, ByVal element As String, ByVal databaseWebID As String, ByVal typeFilter As String, ByVal AFObject As String, ByVal wsName As String, ByVal pointSource As String)

    ' Improve performance by disabling automatic calculation and screen updating
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ActiveSheet.DisplayPageBreaks = False

    ' Variable declaration
    Dim URL As String
    Dim json As Variant
    Dim dict As Object
    Dim items As Collection
    Dim category_concat As String
    Dim ws As Worksheet
    Dim businessArea As String
    Dim dict_point As Dictionary
    Dim columns As New Collection
    Dim i As Integer

    ' Try to get the worksheet with the specified name
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0

    ' If the worksheet does not exist, create a new one
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If

    ' Define columns
    columns.Add "Path"
    columns.Add "Name"
    columns.Add "Element"
    columns.Add "Attribute"
    columns.Add "Description"
    columns.Add "HasChildren"
    columns.Add "TemplateName"
    columns.Add "CategoryNames"
    columns.Add "IsHidden"
    columns.Add "IsExcluded"
    columns.Add "DefaultUnitsNameAbbreviation"
    columns.Add "Type"
    columns.Add "Value"
    columns.Add "DataReferencePlugIn"
    columns.Add "ConfigString"
    columns.Add "Good"
    columns.Add "Timestamp"

    ' Define headers
    Dim header(16) As String
    header(0) = "Parent"
    header(1) = "Name"
    header(2) = "ObjectType"
    header(3) = "Description"
    header(4) = "ReferenceType"
    header(5) = "Template"
    header(6) = "Categories"
    header(7) = "Status"
    header(8) = "AttributeIsHidden"
    header(9) = "AttributeIsExcluded"
    header(10) = "AttributeDefaultUOM"
    header(11) = "AttributeType"
    header(12) = "AttributeValue"
    header(13) = "AttributeDataReference"
    header(14) = "AttributeConfigString"
    header(15) = "Status"
    header(16) = "TimeStamp"

    ' Set worksheet header
    ws.Range("A1:P1").Value = header

    ' Build base URL
    Dim URLbases() As String
    Dim URLbase As String
    URLbases = Split(endpoint, "/")
    For i = 0 To 3
        URLbase = URLbase & URLbases(i) & "/"
    Next i

    ' Configuration of URLs for different AF objects
    Dim keys_Elements As String
    Dim keys_Attributes As String
    Dim keys_Points As String

    keys_Elements = "&selectedFields=Items.Path;Items.Name;Items.Element;Items.Description;Items.TemplateName;Items.HasChildren;Items.CategoryNames;Items.Links.Analyses;Items.Links.Attributes;Items.Links.Elements;Links.First;Items.Status"
    keys_Attributes = "&selectedFields=Items.Name;Items.Description;Items.Path;Items.Type;Items.DataReferencePlugIn;Items.ConfigString;Items.IsExcluded;Items.IsHidden;Items.HasChildren;Items.CategoryNames;Items.Links.Attributes;Items.Links.Value;Items.Links.Point"
    keys_Points = "&selectedFields=Items.Name;Items.Path;Items.PointSource;Items.Descriptor;Items.Links.Value"

    ' Build query URL based on the specified AF object
    Select Case AFObject
        Case "attributes"
            parts = Split(path, "\\")
            businessArea = parts(UBound(parts))
            URL = URLbase & AFObject & "/search?databaseWebID=" & databaseWebID & "&query=" & typeFilter & ":=" & "'" & element & "'" & " Element:{Template:=" & "'" & businessArea & "'" & "}" & "&maxCount=100000" & keys_Attributes
        Case "elements", "analyses"
            URL = URLbase & AFObject & "/search?databaseWebID=" & databaseWebID & "&query=" & typeFilter & ":=" & "'" & element & "'" & "&maxCount=100000" & keys_Elements
        Case "points"
            URL = URLbase & AFObject & "/search?dataServerWebId=" & databaseWebID & "&query=name:=" & "'" & element & "'" & " AND " & pointSource & "&maxCount=100000" & keys_Points
        Case Else
            MsgBox "Error: AFObject not found!"
            Exit Sub
    End Select

    ' Get API response
    json = GETResponseAPI.GetAPIResponse(URL)
    If json(1) <> 200 Then
        MsgBox json(0)
        Exit Sub
    End If

    ' Parse JSON response
    Set dict = JsonConverter.ParseJson(json(0))
    Set items = dict("Items")

    ' Prepare to write data to worksheet
    Dim rows As Long
    Dim Data() As Variant
    ReDim Data(columns.Count)
    Dim lastCell As Range
    Set lastCell = ws.Columns("A").Cells(ws.Rows.Count).End(xlUp)
    rows = lastCell.Row + 1

    ' Extraction modes
    Dim element_mode As Boolean
    Dim attribute_mode As Boolean
    Dim analysis_mode As Boolean
    Dim current_value_mode As Boolean
    Dim points_mode As Boolean

    element_mode = True
    attribute_mode = True
    analysis_mode = True
    current_value_mode = True
    points_mode = True

    ' Process extracted items
    Dim item As Variant
    For Each item In items
        If element_mode Then ProcessElement item, ws, Data, rows
        If attribute_mode Then ProcessAttribute item, ws, Data, rows, current_value_mode
        If analysis_mode Then ProcessAnalysis item, ws, Data, rows
        If points_mode Then ProcessPoints item, ws, Data, rows
    Next item

    ' Restore performance settings
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ActiveSheet.DisplayPageBreaks = True

    ' Auto fit columns and finish
    ws.Columns.AutoFit
    MsgBox "Processing completed. Rows processed: " & rows

End Sub

' Function to process elements
Sub ProcessElement(item As Variant, ws As Worksheet, Data() As Variant, ByRef rows As Long)
    ' Extract and save element data
    ' Implement element processing logic
End Sub

' Function to process attributes
Sub ProcessAttribute(item As Variant, ws As Worksheet, Data() As Variant, ByRef rows As Long, current_value_mode As Boolean)
    ' Extract and save attribute data
    ' Implement attribute processing logic
End Sub

' Function to process analyses
Sub ProcessAnalysis(item As Variant, ws As Worksheet, Data() As Variant, ByRef rows As Long)
    ' Extract and save analysis data
    ' Implement analysis processing logic
End Sub

' Function to process points
Sub ProcessPoints(item As Variant, ws As Worksheet, Data() As Variant, ByRef rows As Long)
    ' Extract and save point data
    ' Implement point processing logic
End Sub

' Auxiliary functions
Function ElementExistAtCollection(ByVal collection As Collection, ByVal element As Variant) As Boolean
    Dim i As Integer
    For i = 1 To collection.Count
        If collection.Item(i) = element Then
            ElementExistAtCollection = True
            Exit Function
        End If
    Next i
    ElementExistAtCollection = False
End Function

Function CutPath(longPath As String) As String
    Dim index As Integer
    Dim cutPath As String
    Dim parts As Variant
    Dim newPath As String

    index = InStr(longPath, "?path=")
    cutPath = Mid(longPath, index + 6)
    parts = Split(cutPath, "\")
    ReDim Preserve parts(UBound(parts) - 1)
    newPath = Join(parts, "\")
    CutPath = Replace(newPath, "\\", "\")
End Function

Function checkPathTarget(path As String, pathTarget As String) As Boolean
    checkPathTarget = (InStr(pathTarget, path) > 0)
End Function
