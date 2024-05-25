' Reference: Microsoft Scripting Runtime
' Reference: Microsoft WinHTTP Services, version 5.1
' To use module JsonConverter, you must add a reference to the Microsoft Scripting Runtime library: https://github.com/VBA-tools/VBA-JSON

Sub ExtractPIWebAPI(endpoint As String)

    ' Disable Excel features to optimize script performance
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ActiveSheet.DisplayPageBreaks = False

    ' Declare global variables
    Dim json As String
    Dim dict As Object
    Dim dict_elements As Object
    Dim dict_value As Dictionary
    Dim Attributes As String
    Dim Elements As String
    Dim Items As Collection
    Dim Link As Dictionary
    Dim Items_Attributes As Collection
    Dim Items_children As Collection
    Dim Items_Elements As Collection
    Dim Links_Attributes As Dictionary
    Dim Links_Children As Dictionary
    Dim CategoryNames As Collection
    Dim category As Variant
    Dim category_concat As String
    Dim ws As Worksheet

    ' Check if a worksheet named "PI Tags" exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PI Tags")
    On Error GoTo 0

    ' If the worksheet does not exist, create a new one
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "PI Tags"
    Else
        ws.Cells.Clear
    End If

    Dim columns As New Collection
    Dim i As Integer

    ' Create an enumeration of references from the API
    columns.Add "Path" ' A - 1
    columns.Add "Name" ' B - 2
    columns.Add "Element" ' C - 3
    columns.Add "Attribute" ' C - 4
    columns.Add "Description" ' D - 5
    columns.Add "HasChildren" ' E - 6
    columns.Add "TemplateName" ' F - 7
    columns.Add "CategoryNames" ' G - 8
    columns.Add "IsHidden" ' H - 9
    columns.Add "IsExcluded" ' I - 10
    columns.Add "DefaultUnitsNameAbbreviation" ' J - 11
    columns.Add "Type" ' K - 12
    columns.Add "Value" ' L - 13
    columns.Add "DataReferencePlugIn" ' M - 14
    columns.Add "ConfigString" ' N - 15
    columns.Add "Good" ' O - 16
    columns.Add "Timestamp" ' P - 17

    ' Create the header in the "PI Tags" worksheet similar to PI Builder
    Dim header(16) As String
    header(0) = "Parent"
    header(1) = "Name"
    header(2) = "ObjectType"
    header(3) = "Description"
    header(4) = "ReferenceType"
    header(5) = "Template"
    header(6) = "Categories"
    header(7) = "AttributeIsHidden"
    header(8) = "AttributeIsExcluded"
    header(9) = "AttributeDefaultUOM"
    header(10) = "AttributeType"
    header(11) = "AttributeValue"
    header(12) = "AttributeDataReference"
    header(13) = "AttributeConfigString"
    header(14) = "Status"
    header(15) = "TimeStamp"

    ws.Range("A1:P1").Value = header ' Save the header data in the worksheet at row 1

    ' GET HTTP of the first element in the hierarchy (TOP-DOWN) to be explored
    json = GetAPIResponse(endpoint)

    Set dict = JsonConverter.ParseJson(json)
    Set Items = dict("Items")
    Set Link = dict("Links")
    Dim Link_return As String
    Link_return = Link("First")

    ' Create variables to store the hierarchy data based on nodes and branches
    Dim nodes() As Integer
    Dim branch() As Integer
    Dim Link_tree() As String
    Dim Data() As Variant ' Data to be saved in the worksheet
    ReDim Data(columns.Count)

    Dim count_branch As Integer
    Dim aux_branch As Integer
    Dim count_node As Integer
    Dim rows As Integer

    ' Initialize hierarchy control variables
    count_branch = 1
    count_node = 0
    aux_branch = 1
    rows = 2

    ReDim nodes(count_node)
    nodes(count_node) = Items.Count
    ReDim Link_tree(count_node)
    ReDim Preserve branch(count_node)
    branch(count_node) = aux_branch
    Link_tree(count_node) = Link_return

    Dim exit_hierarchy As Boolean
    Dim element_mode As Boolean
    Dim attribute_mode As Boolean
    Dim children_mode As Boolean

    ' Initialize variables to export to the "PI Tags" worksheet
    exit_hierarchy = True
    element_mode = True
    attribute_mode = True
    children_mode = True

    While exit_hierarchy ' Loop to traverse the element hierarchy

        ' Extract Elements of tree PI AF
        If element_mode = True Then
            Set CategoryNames = Items(aux_branch)("CategoryNames")
            category_concat = ""
            For Each category In CategoryNames
                category_concat = category_concat & category & ";"
            Next category

            Data(0) = Items(aux_branch)("Path")
            Data(1) = Items(aux_branch)("Name") ' Name
            Data(2) = columns(3) ' Element
            Data(3) = Items(aux_branch)("Description") ' Description
            Data(4) = "Parent-Child" ' ReferenceType
            Data(5) = Items(aux_branch)("TemplateName") ' TemplateName
            Data(6) = category_concat ' CategoryNames

            ws.Range("A" & rows & ":G" & rows).Value = Data ' Save the data in the worksheet
            rows = rows + 1 ' Increment the row to save the data in the worksheet
        End If

        ' Extract Attributes of tree PI AF
        If attribute_mode = True Then
            Attributes = Items(aux_branch)("Links")("Attributes")
            Set dict = JsonConverter.ParseJson(GetAPIResponse(Attributes))
            Set Items_Attributes = dict("Items")

            Dim Item_Attribute As Variant
            For Each Item_Attribute In Items_Attributes
                Set Links_Attributes = Item_Attribute("Links")
                Set CategoryNames = Item_Attribute(columns(8))
                category_concat = ""
                For Each category In CategoryNames
                    category_concat = category_concat & category & ";"
                Next category

                Data(0) = Items(aux_branch)("Path") ' Parent
                Data(1) = Item_Attribute(columns(2)) ' Name
                Data(2) = columns(4) ' Attribute
                Data(3) = Item_Attribute(columns(5)) ' Description
                Data(4) = "" ' ReferenceType is empty for Attribute and Analyses
                Data(5) = Item_Attribute(columns(7)) ' TemplateName
                Data(6) = category_concat ' CategoryNames
                Data(7) = Item_Attribute(columns(9)) ' IsHidden
                Data(8) = Item_Attribute(columns(10)) ' IsExcluded
                Data(9) = Item_Attribute(columns(11)) ' DefaultUnitsNameAbbreviation
                Data(10) = Item_Attribute(columns(12)) ' Type
                Data(11) = "" ' Value is empty for Attribute and Analyses
                Data(12) = Item_Attribute(columns(14)) ' DataReferencePlugIn
                Data(13) = Item_Attribute(columns(15)) ' ConfigString

                ws.Range("A" & rows & ":P" & rows).Value = Data ' Save the data in the worksheet
                rows = rows + 1 ' Increment the row to save the data in the worksheet

                ' Extract Attributes child of tree PI AF
                If children_mode = True Then
                    If Item_Attribute("HasChildren") = True Then
                        Set dict_children = JsonConverter.ParseJson(GetAPIResponse(Links_Attributes("Attributes")))
                        Set Items_children = dict_children("Items")
                        Dim Item_children As Variant
                        For Each Item_children In Items_children
                            Set Links_Children = Item_children("Links")
                            Set CategoryNames = Item_children(columns(8))
                            category_concat = ""
                            For Each category In CategoryNames
                                category_concat = category_concat & category & ";"
                            Next category

                            Data(0) = Items(aux_branch)("Path") ' Parent
                            Data(1) = Item_Attribute(columns(2)) & "|" & Item_children(columns(2)) ' Name
                            Data(2) = columns(4) ' Attribute
                            Data(3) = Item_children(columns(5)) ' Description
                            Data(4) = "" ' ReferenceType is empty for Attribute and Analyses
                            Data(5) = Item_children(columns(7)) ' TemplateName
                            Data(6) = category_concat ' CategoryNames
                            Data(7) = Item_children(columns(9)) ' IsHidden
                            Data(8) = Item_children(columns(10)) ' IsExcluded
                            Data(9) = Item_children(columns(11)) ' DefaultUnitsNameAbbreviation
                            Data(10) = Item_children(columns(12)) ' Type
                            Data(11) = "" ' Value is empty for Attribute and Analyses
                            Data(12) = "" ' DataReferencePlugIn
                            Data(13) = "" ' ConfigString

                            ws.Range("A" & rows & ":P" & rows).Value = Data ' Save the data in the worksheet
                            rows = rows + 1 ' Increment the row to save the data in the worksheet
                        Next Item_children
                    End If
                End If
            Next Item_Attribute
        End If

        ' Extract Analyses of tree PI AF

        ' Tree search conditions
        If Items(aux_branch)("HasChildren") = True Then ' Check if the element has children, if so, enter the loop to traverse the children
            Elements = Items(aux_branch)("Links")("Elements")
            Set dict_elements = JsonConverter.ParseJson(GetAPIResponse(Elements))
            Set Items_Elements = dict_elements("Items")

            aux_branch = 1
            count_branch = 1
            ReDim Preserve nodes(UBound(nodes) + 1)
            nodes(UBound(nodes)) = Items_Elements.Count
            count_node = Items_Elements.Count

            ReDim Preserve branch(UBound(branch) + 1)
            branch(UBound(branch)) = count_branch

            Set Items = Items_Elements

            ReDim Preserve Link_tree(UBound(Link_tree) + 1)
            Link_tree(UBound(Link_tree)) = dict_elements("Links")("First") ' Save the link of the first node of the tree
        ElseIf Items(aux_branch)("HasChildren") = False And count_node > 1 Then ' Check if the elements have "siblings", if so, enter the loop to traverse the "siblings"
            aux_branch = aux_branch + 1
            count_branch = count_branch + 1
            branch(UBound(branch)) = count_branch
            count_node = count_node - 1
        ElseIf Items(aux_branch)("HasChildren") = False And count_node <= 1 Then ' Check if the element has no children and no "siblings", if so, go back to the loop to traverse the "parents"
            ' This condition is important for controlling which nodes and branches have already been traversed
            If AreArraysEqual(nodes, branch) = True Then
                exit_hierarchy = False ' Exit the hierarchy while loop
            Else
                Dim check_size_array As Boolean
                check_size_array = False

                ' Check if the size of the array is zero
                While Not (check_size_array)
                    If branch(UBound(branch)) = nodes(UBound(nodes)) Then
                        ReDim Preserve branch(0 To UBound(branch) - 1)
                        ReDim Preserve nodes(0 To UBound(nodes) - 1)
                        ReDim Preserve Link_tree(0 To UBound(Link_tree) - 1)
                        check_size_array = False
                    Else
                        check_size_array = True
                    End If
                Wend

                branch(UBound(branch)) = branch(UBound(branch)) + 1

                ' End of verification of which nodes and branches have already been traversed and control array adjustments
                aux_branch = branch(UBound(branch))

                Set dict = JsonConverter.ParseJson(GetAPIResponse(Link_tree(UBound(Link_tree))))
                Set Items = dict("Items")
                Link_tree(UBound(Link_tree)) = dict("Links")("First")
                count_branch = 1
                count_node = 0
            End If
        End If

    Wend

    ' Enable Excel features
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ActiveSheet.DisplayPageBreaks = True

    ws.Columns.AutoFit ' Adjust column width
End Sub

Function Base64Encode(ByVal sText As String) As String
    Dim oXML As Object
    Dim oNode As Object
    Set oXML = CreateObject("Msxml2.DOMDocument")
    Set oNode = oXML.createElement("base64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.Text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Stream_StringToBinary(ByVal sText As String) As Variant
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "us-ascii"
    ado.Open
    ado.WriteText sText
    ado.Position = 0
    ado.Type = 1
    Stream_StringToBinary = ado.Read
End Function

Function GetAPIResponse(ByVal url As String) As String
    ' Anonymized username and password
    Dim username As String
    Dim password As String
    username = "your_username"
    password = "your_password"

    ' This function returns the response from a URL using Basic Authentication
    Dim xmlHttp As Object
    Dim response As String

    ' Create an instance of the XMLHTTP object
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")

    ' Open a connection to the API endpoint with Basic Authentication
    xmlHttp.Open "GET", url, False
    xmlHttp.SetRequestHeader "Authorization", "Basic " & Base64Encode(username & password)
    ' Send the HTTP request and retrieve the response
    xmlHttp.Send
    ' Return the API response
    GetAPIResponse = xmlHttp.ResponseText
End Function
