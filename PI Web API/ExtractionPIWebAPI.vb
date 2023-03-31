' Referência: Microsoft Scripting Runtime
' Referência: Microsoft WinHTTP Services, versão 5.1
' To use module JsonConverter, you must add a reference to the Microsoft Scripting Runtime library: https://github.com/VBA-tools/VBA-JSON


Sub ExtractPIWebAPI(url As String)

'OptimizeVBA (False)
Application.Calculation = IIf(False, xlCalculationManual, xlCalculationAutomatic)
Application.EnableEvents = Not (False)
Application.ScreenUpdating = Not (False)
ActiveSheet.DisplayPageBreaks = Not (False)
Application.EnableEvents = False

time1 = Timer


'Dim url As String
Dim json As String
Dim dict As Object
Dim dict_analyses As Object
Dim dict_elements As Object
Dim dict_value As Dictionary
Dim user As String
Dim pass As String
Dim Attributes As String
Dim Analysis As String
Dim Elements As String

Dim Items As Collection
Dim Link As Dictionary
Dim Items_Attributes As Collection
Dim Items_children As Collection
Dim Items_Analyses As Collection
Dim Items_Elements As Collection
Dim Items_Value As String
Dim Links_Attributes As Dictionary
Dim Links_Children As Dictionary
Dim Links_Analysis As Dictionary
Dim Links As Dictionary
Dim CategoryNames As Collection
Dim category As Variant
Dim Value As String
Dim Link_return As String
Dim dict_children As Dictionary
Dim Name As Dictionary
Dim category_concat As String
Dim ws As Worksheet
Dim TagGood As Integer
Dim TagBad As Integer


Set ws = ThisWorkbook.Worksheets("ExtractPIWebAPI")


Dim columns As New Collection
Dim i As Integer
    
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
columns.Add "DataReferencePlugIn" ' M -  14
columns.Add "ConfigString" ' N - 15
columns.Add "Good" ' O - 16
columns.Add "Timestamp" ' P - 17

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
header(14) = "Value"
header(15) = "TimeStamp"

ws.Range("A" & 1 & ":P" & 1).Value = header ' salva os dados de cabeçalho na planilha na linha 1



'url = "https://af-dev-ho.oxy.com/piwebapi/elements/F1EmQ7Uvuf0ZDkqSRWtZTQP2bQ9ptzWi1A7RGROQBQVqVzmgQUYtREVWLUhPLk9YWS5DT01cREVWX09SQ00gU09MVVRJT05TXFBFUk1JQU4gRU9SXEhPQkJTXEZBQ0lMSVRJRVM/elements"
 
json = GetAPIResponse(url)
   
    
Set dict = JsonConverter.ParseJson(json)
Set Items = dict("Items")
Set Link = dict("Links")
Link_return = Link("First")

Dim nodes() As Integer '
Dim branch() As Integer '
Dim Link_tree() As String '
Dim data() As Variant
ReDim data(columns.Count)


Dim count_branch As Integer
Dim aux_branch As Integer
Dim count_node As Integer
Dim rows As Integer



count_branch = 1
count_node = 0
aux_branch = 1
rows = 2

TagGood = 0
TagBad = 0

ReDim nodes(count_node)
nodes(count_node) = Items.Count
ReDim Link_tree(count_node)
ReDim Preserve branch(count_node)
branch(count_node) = aux_branch
Link_tree(count_node) = Link_return


Dim exit_overall As Boolean
exit_overall = True

Dim element_mode As Boolean
Dim attribute_mode As Boolean
Dim analysis_mode As Boolean
Dim children_mode As Boolean

element_mode = configAPI.cbx_Elements.Value 'False
attribute_mode = configAPI.cbx_Attributes 'True
analysis_mode = configAPI.cbx_analysis 'False
children_mode = configAPI.cbx_child_attribute 'True

While exit_overall
    
    Set Links = Items(aux_branch)("Links")

    Elements = Links("Elements")
    Set dict_elements = JsonConverter.ParseJson(GetAPIResponse(Elements))
    Set Items_Elements = dict_elements("Items")
    

    ' ######## Extract Elements of tree PI AF ###############################

    If element_mode = True Then
        
        Set CategoryNames = Items(aux_branch)("CategoryNames")
        category_concat = ""
        For Each category In CategoryNames
                category_concat = category_concat & category & ";"
            Next category

        data(0) = Items(aux_branch)("Path")
        data(1) = Items(aux_branch)("Name") ' Name
        data(2) = columns(3) ' Element
        data(3) = Items(aux_branch)("Description") ' Description
        data(4) = "Parent-Child" ' ReferenceType
        data(5) = Items(aux_branch)("TemplateName") ' TemplateName
        data(6) = category_concat '  CategoryNames

        ws.Range("A" & rows & ":G" & rows).Value = data ' salva os dados na planilha
    
        rows = rows + 1 ' incrementa a linha para salvar os dados na planilha

    End If

   ' ######## Extract Attributes of tree PI AF ################################

   If attribute_mode = True Then
        Attributes = Links("Attributes")
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
            
            Set dict_value = JsonConverter.ParseJson(GetAPIResponse(Links_Attributes("Value")))
            
            If VarType(dict_value("Value")) = 9 Then
                Set Name = dict_value("Value")
                Items_Value = Name("Name")
            Else
                Items_Value = dict_value("Value")
            End If

            data(0) = Items(aux_branch)("Path") ' Parent
            data(1) = Item_Attribute(columns(2)) ' Name
            data(2) = columns(4) ' Attribute
            data(3) = Item_Attribute(columns(5)) ' Description
            data(4) = "" ' ReferenceType is empty for Attribute and Analyses
            data(5) = Item_Attribute(columns(7)) ' TemplateName
            data(6) = category_concat  '  CategoryNames
            data(7) = Item_Attribute(columns(9)) ' IsHidden
            data(8) = Item_Attribute(columns(10)) ' IsExcluded
            data(9) = Item_Attribute(columns(11)) ' DefaultUnitsNameAbbreviation
            data(10) = Item_Attribute(columns(12)) ' Type
            data(11) = Items_Value ' * Value
            data(12) = Item_Attribute(columns(14)) ' DataReferencePlugIn
            data(13) = Item_Attribute(columns(15)) ' ConfigString
            
            If dict_value(columns(16)) = "True" Then
                data(14) = "Good"
                TagGood = TagGood + 1
            Else
                data(14) = "Bad"
                TagBad = TagBad + 1
            End If
            
            'data(14) = IIf(dict_value(columns(16)) = "True", "Good", "Bad") ' Good
            data(15) = dict_value(columns(17)) ' Timestamp
        
            ws.Range("A" & rows & ":P" & rows).Value = data ' salva os dados na planilha
            rows = rows + 1 ' incrementa a linha para salvar os dados na planilha
        
        
            ' ######## Extract Attributes child  of tree PI AF ################################
            If children_mode = True Then
                Set dict_children = JsonConverter.ParseJson(GetAPIResponse(Links_Attributes("Attributes")))
                
                If dict_children("Items").Count <> 0 Then
                    
                    Set Items_children = dict_children("Items")
                    
                    Dim Item_children As Variant
                    For Each Item_children In Items_children
                    
                        Set Links_Children = Item_children("Links")
                    
                        Set CategoryNames = Item_children(columns(8))
                        category_concat = ""
                        For Each category In CategoryNames
                            category_concat = category_concat & category & ";"
                        Next category
                        
                        Set dict_value = JsonConverter.ParseJson(GetAPIResponse(Links_Children("Value")))
                
                        If VarType(dict_value("Value")) = 9 Then
                            Set Name = dict_value("Value")
                            Items_Value = Name("Name")
                        
                        Else
                            Items_Value = dict_value("Value")
                        
                        End If
                        
                        data(0) = Items(aux_branch)("Path") ' Parent
                        data(1) = Item_Attribute(columns(2)) & "|" & Item_children(columns(2)) ' Name
                        data(2) = columns(4) ' Attribute
                        data(3) = Item_children(columns(5)) ' Description
                        data(4) = "" ' ReferenceType is empty for Attribute and Analyses
                        data(5) = Item_children(columns(7)) ' TemplateName
                        data(6) = category_concat  '  CategoryNames
                        data(7) = Item_children(columns(9)) ' IsHidden
                        data(8) = Item_children(columns(10)) ' IsExcluded
                        data(9) = Item_children(columns(11)) ' DefaultUnitsNameAbbreviation
                        data(10) = Item_children(columns(12)) ' Type
                        data(11) = Items_Value ' Value
                        data(12) = "" ' DataReferencePlugIn
                        data(13) = "" 'ConfigString
                        If dict_value(columns(16)) = "True" Then
                            data(14) = "Good"
                            TagGood = TagGood + 1
                        Else
                            data(14) = "Bad"
                            TagBad = TagBad + 1
                        End If
                        'data(14) = IIf(dict_value(columns(16)) = "True", "Good", "Bad") ' Good
                        data(15) = dict_value(columns(17)) ' Timestamp
                            
                        ws.Range("A" & rows & ":P" & rows).Value = data ' salva os dados na planilha
                    
                        rows = rows + 1 ' incrementa a linha para salvar os dados na planilha

                    Next Item_children
                End If
            
            End If
        Next Item_Attribute
    End If

    ' ######## Extract Analyses of tree PI AF ################################

    If analysis_mode = True Then
            Analysis = Links("Analyses")
            Set dict_analyses = JsonConverter.ParseJson(GetAPIResponse(Analysis))
            Set Items_Analyses = dict_analyses("Items")
            
            Dim Item_Analysis As Variant
            For Each Item_Analysis In Items_Analyses
            
            
                Set CategoryNames = Item_Analysis(columns(8))
                category_concat = ""
                For Each category In CategoryNames
                category_concat = category_concat & category & ";"
                Next category
                
            data(0) = Items(aux_branch)("Path")
            data(1) = Item_Analysis(columns(2)) ' Name
            data(2) = "Analysis" ' Analysis
            data(3) = Item_Analysis(columns(5)) ' Description
            data(4) = "" ' ReferenceType is empty for Attribute and Analyses
            data(5) = Item_Analysis(columns(7)) ' TemplateName
            data(6) = category_concat  '  CategoryNames
            
            ws.Range("A" & rows & ":G" & rows).Value = data ' salva os dados na planilha
                
                rows = rows + 1 ' incrementa a linha para salvar os dados na planilha
                
            Next
    End If

    ' ######## Tree search conditions  ################################
    
    If Items_Elements.Count >= 1 Then ' Into in each first node of the tree
    
            aux_branch = 1
            ReDim Preserve nodes(UBound(nodes) + 1)
            nodes(UBound(nodes)) = Items_Elements.Count
            count_node = Items_Elements.Count

            ReDim Preserve branch(UBound(branch) + 1) '
            branch(UBound(branch)) = count_branch '

            Set dict = JsonConverter.ParseJson(GetAPIResponse(Elements))
            Set Items = dict("Items")
            Set Link = dict("Links")
            Link_return = Link("First")
            
            ReDim Preserve Link_tree(UBound(Link_tree) + 1)
            Link_tree(UBound(Link_tree)) = Link_return
            
            
    ElseIf Items_Elements.Count < 1 And count_node > 1 Then

            aux_branch = aux_branch + 1
            count_branch = count_branch + 1
            
            branch(UBound(branch)) = count_branch
            
            count_node = count_node - 1
            

    ElseIf Items_Elements.Count < 1 And count_node <= 1 Then '
        
        If nodes(0) = branch(0) And nodes(1) = branch(1) Then
            
            exit_overall = False
        Else
            If count_node = 0 Then
                    branch(UBound(branch)) = branch(UBound(branch)) + 1
            Else
                    ReDim Preserve branch(0 To UBound(branch) - 1)
                    ReDim Preserve nodes(0 To UBound(nodes) - 1)
                    ReDim Preserve Link_tree(0 To UBound(Link_tree) - 1)
                    
                    
                    Dim test As Boolean
                    test = False
                    
                    While Not (test):
                        If branch(UBound(branch)) = nodes(UBound(nodes)) Then '
                        
                            ReDim Preserve branch(0 To UBound(branch) - 1)
                            ReDim Preserve nodes(0 To UBound(nodes) - 1)
                            ReDim Preserve Link_tree(0 To UBound(Link_tree) - 1)
                            test = False
                        
                        Else
                            test = True
                        
                        End If
                    Wend
                    
                    branch(UBound(branch)) = branch(UBound(branch)) + 1
            End If
            
            
        aux_branch = branch(UBound(branch))
                
            
        Set dict = JsonConverter.ParseJson(GetAPIResponse(Link_tree(UBound(Link_tree))))
        
        
        Set Items = dict("Items")
        Set Link = dict("Links")
        Link_return = Link("First")
        
        Link_tree(UBound(Link_tree)) = Link_return
        
        count_branch = 1
        count_node = 0
                
            
        End If
                

    End If
   
    
Wend



'OptimizeVBA (True)
Application.Calculation = IIf(True, xlCalculationManual, xlCalculationAutomatic)
Application.EnableEvents = Not (True)
Application.ScreenUpdating = Not (True)
ActiveSheet.DisplayPageBreaks = Not (True)
Application.EnableEvents = True

popup_finish.text_attributeGoodTotal.Value = TagGood
popup_finish.text_attributeBadTotal.Value = TagBad
popup_finish.text_runtime.Value = CStr((Timer - time1) / 60) & " min"
popup_finish.text_rowsTotal.Value = rows

popup_finish.Show
Debug.Print "Finish row: " & rows & " Runtime: " & CStr((Timer - time1) / 60) & " min"
MsgBox "Finish row: " & rows & " Runtime: " & CStr(Timer - time1)




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
  Dim xmlHttp As Object
  ' Dim url As String
  Dim username As String
  Dim password As String
  Dim response As String
  
  ' Define the API endpoint URL
  ' url = "https://af-houston.oxy.com/piwebapi/"
  'url = "https://af-dev-ho.oxy.com/piwebapi/elements/F1EmQ7Uvuf0ZDkqSRWtZTQP2bQ85tzWi1A7RGROQBQVqVzmgQUYtREVWLUhPLk9YWS5DT01cREVWX09SQ00gU09MVVRJT05TXFBFUk1JQU4gRU9SXEhPQkJT/elements"
  
  ' Define the Basic Authentication credentials
  username = "kaike_carvalho@oxy.com"
  password = ""
  
  ' Create an instance of the XMLHTTP object
  Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
  
  ' Open a connection to the API endpoint with Basic Authentication
  xmlHttp.Open "GET", url, False
  xmlHttp.SetRequestHeader "Authorization", "Basic " & Base64Encode(username & ":" & password)
  
  ' Send the HTTP request and retrieve the response
  xmlHttp.Send
  
      
  ' Return the API response
  GetAPIResponse = xmlHttp.ResponseText
  
  
End Function









