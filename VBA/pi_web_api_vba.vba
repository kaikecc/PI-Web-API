' Referência: Microsoft Scripting Runtime
' Referência: Microsoft WinHTTP Services, versão 5.1
' To use module JsonConverter, you must add a reference to the Microsoft Scripting Runtime library: https://github.com/VBA-tools/VBA-JSON


Sub ExtractPIWebAPI(endpoint As String)

    ' Desabilita recursos do Excel para otimizar o desempenho da execução do script
    Application.Calculation = IIf(False, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (False)
    Application.ScreenUpdating = Not (False)
    ActiveSheet.DisplayPageBreaks = Not (False)
    Application.EnableEvents = False

    ' Cria variáveis globais 
    Dim json As String
    Dim dict As Object
    Dim dict_analyses As Object
    Dim dict_elements As Object
    Dim dict_value As Dictionary
    Dim Attributes As String
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

    ' Verifica se existe uma aba chamada PI Tags
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PI Tags")
    On Error GoTo 0

    ' Se a aba não existor, então criar uma nova
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "PI Tags"   
    Else
        ws.Cells.Clear
    End If

    Dim columns As New Collection
    Dim i As Integer

    ' Cria um enumateset de referências provinientes da API    
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

    ' Cria o cabeçalho na planilha PI Tags semelhante ao PI Builder
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


    ws.Range("A" & 1 & ":P" & 1).Value = header ' salva os dados de cabeçalho na planilha na linha 1

    ' GET HTTP do primeiro elemento da hierarquia(TOP-DOWN) a ser explorada 
    json = GetAPIResponse(endpoint)   
    
    Set dict = JsonConverter.ParseJson(json)
    Set Items = dict("Items")
    Set Link = dict("Links")
    Link_return = Link("First")

    ' Cria variáveis para armazenar os dados da hierarquia baseado em Nós e Ramificações                        
    Dim nodes() As Integer '
    Dim branch() As Integer '
    Dim Link_tree() As String '
    Dim Data() As Variant ' Dados a serem salvos na planilha
    ReDim Data(columns.Count)

    Dim count_branch As Integer
    Dim aux_branch As Integer
    Dim count_node As Integer
    Dim rows As Integer

    ' Inicializa as variáveis de controle da hierarquia
    count_branch = 1
    count_node = 0
    aux_branch = 1
    rows = 2

    ' Inicializa as variáveis de controle da hierarquia
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

    ' Inicializa as variáveis para exportar na planilha PI Tags
    exit_hierarchy = True
    element_mode = True
    attribute_mode = True
    children_mode = True

    While exit_hierarchy ' Loop para percorrer a hierarquia de elementos
        
        ' ######## Extract Elements of tree PI AF ###############################

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
            Data(6) = category_concat '  CategoryNames

            ws.Range("A" & rows & ":G" & rows).Value = Data ' salva os dados na planilha    
            rows = rows + 1 ' incrementa a linha para salvar os dados na planilha
        End If

        ' ######## Extract Attributes of tree PI AF ################################

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
                    Data(6) = category_concat  '  CategoryNames
                    Data(7) = Item_Attribute(columns(9)) ' IsHidden
                    Data(8) = Item_Attribute(columns(10)) ' IsExcluded
                    Data(9) = Item_Attribute(columns(11)) ' DefaultUnitsNameAbbreviation
                    Data(10) = Item_Attribute(columns(12)) ' Type
                    Data(11) = "" ' Value is empty for Attribute and Analyses
                    Data(12) = Item_Attribute(columns(14)) ' DataReferencePlugIn
                    Data(13) = Item_Attribute(columns(15)) ' ConfigString
                    
                    ws.Range("A" & rows & ":P" & rows).Value = Data ' salva os dados na planilha
                    rows = rows + 1 ' incrementa a linha para salvar os dados na planilha        
                
                    ' ######## Extract Attributes child  of tree PI AF ################################
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
                                Data(6) = category_concat  '  CategoryNames
                                Data(7) = Item_children(columns(9)) ' IsHidden
                                Data(8) = Item_children(columns(10)) ' IsExcluded
                                Data(9) = Item_children(columns(11)) ' DefaultUnitsNameAbbreviation
                                Data(10) = Item_children(columns(12)) ' Type  
                                Data(11) = ""  ' Value is empty for Attribute and Analyses                   
                                Data(12) = "" ' DataReferencePlugIn
                                Data(13) = "" 'ConfigString                       
                                    
                                ws.Range("A" & rows & ":P" & rows).Value = Data ' salva os dados na planilha                    
                                rows = rows + 1 ' incrementa a linha para salvar os dados na planilha
                            Next Item_children
                        End If            
                    End If
                Next Item_Attribute
            End If

        ' ######## Extract Analyses of tree PI AF ################################

        ' ######## Tree search conditions  #######################################
        
        If Items(aux_branch)("HasChildren") = True Then ' Verifica se o elemento tem filhos, se sim entra no loop para percorrer os filhos

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
                Link_tree(UBound(Link_tree)) = dict_elements("Links")("First") ' save the link of the first node of the tree            
                
        ElseIf Items(aux_branch)("HasChildren") = False And count_node > 1 Then ' Verifica se o elementos têm "irmãos", se sim entra no loop para percorrer os "irmãos"

                aux_branch = aux_branch + 1
                count_branch = count_branch + 1            
                branch(UBound(branch)) = count_branch            
                count_node = count_node - 1            

        ElseIf Items(aux_branch)("HasChildren") = False And count_node <= 1 Then ' Verifica se o elemento não tem filhos e não tem "irmãos", se sim volta no loop para percorrer os "pais"
            
            ' Essa condição é importante para o controle de quais nós e ramos já foram percorridos
            If AreArraysEqual(nodes, branch) = True Then           
                exit_hierarchy = False ' sai do while loop da hierarquia
            Else
               
                                               
                Dim check_size_array As Boolean
                check_size_array = False
                
                ' Check if the size of the array is zero
                While Not (check_size_array):
                    If branch(UBound(branch)) = nodes(UBound(nodes)) Then '
                    
                        ReDim Preserve branch(0 To UBound(branch) - 1)
                        ReDim Preserve nodes(0 To UBound(nodes) - 1)
                        ReDim Preserve Link_tree(0 To UBound(Link_tree) - 1)
                        check_size_array = False                        
                    Else
                        check_size_array = True                        
                    End If
                Wend
                
                branch(UBound(branch)) = branch(UBound(branch)) + 1
        
                ' Fim da verificação de quais nós e ramos já foram percorridos e ajustes dos array de controle          
                    
                aux_branch = branch(UBound(branch))              
                    
                Set dict = JsonConverter.ParseJson(GetAPIResponse(Link_tree(UBound(Link_tree))))        
                Set Items = dict("Items")
                Link_tree(UBound(Link_tree)) = dict("Links")("First")        
                count_branch = 1
                count_node = 0             
            End If 
        End If 
         
    Wend

    ' Habilitação dos recursos do Excel
    Application.Calculation = IIf(True, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (True)
    Application.ScreenUpdating = Not (True)
    ActiveSheet.DisplayPageBreaks = Not (True)
    Application.EnableEvents = True

    ws.columns.AutoFit ' Ajusta o tamanho das colunas
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

    Dim username As String
    Dim password As String
    username = ""
    password = ""
 
 ' This function returns the response from a URL using Basic Authentication
  Dim xmlHttp As Object 
  Dim response As String 
   
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











