VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryStringParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' VBA-QueryStringParameters V1.0 alpha
'
' Form Query String for VBA
'
' @class     QueryStringParameters
' @author    Felipe Gustavo<felipegustavo1605@gmail.com>
'
' @depencies VBA ^7
' @depencies Microsoft Scripting Runtime
' ===================================================
'

''
' Code Thows
Public Enum QueryStringParametersThrowsCode
    
    '' Invalid query string syntax
    invalidQueryStringParameters = 3589
    
    '' Empty field in query string
    emptyField = 3590
    
    '' Unknow error happened
    unknowError = 3591
    
    '' Invalid value to assing a new key
    invalidValueToAssing = 3592
    
    '' Value already exists
    valueAlreadyExists = 3593
    
End Enum

''
' This class name
' @var string
Const className As String = "QueryStringParameters"

''
' Dictionary class name
' @var string
Const classNameDictionary As String = "Dictionary"

''
' Represent a Sequential Value
' @var String
Const SEQUENTIAL_KEY = "____sequential_value____"

''
' Code to add how first value
' @var String
Public addAsFirstValue As String

''
' Code to add how last value
' @var String
Public addAsLastValue As String

''
' Save parsed string query in object Collection
' @var \Dictionary
Private parse As New dictionary

''
' Regex Object
' @var \Regex
Private Regex As Object

''
' DOM Object
' @var \HtmlFiles
Private Htmlfile As Object

''
' Make a throw Err.Raise
' @param     {\throws::} Code        Code of Raise
' @param     {string}    ?subMessage Sub message to show
Private Function throw(ByVal Code As QueryStringParametersThrowsCode, Optional ByVal subMessage As String = "")
    
    Dim message As String
    
    newLine = vbCrLf & vbCrLf
        
    Select Case Code
    
    Case QueryStringParametersThrowsCode.invalidQueryStringParameters
        message = "Invalid syntax Query String Parameters" & newLine & subMessage
            
    Case QueryStringParametersThrowsCode.emptyField
        message = "In your query String have a empty field" & newLine & subMessage
            
    Case QueryStringParametersThrowsCode.unknowError
        message = "Unknow error heppened" & newLine & subMessage
        
    Case QueryStringParametersThrowsCode.invalidValueToAssing
        message = "Invalid type var" & newLine & subMessage
    Case QueryStringParametersThrowsCode.valueAlreadyExists
        message = "Value already exists" & newLine & subMessage
    Case Else
        message = subMessage
    
    End Select
    
    Err.Raise Code, className, message
    
End Function

''
' Parsed a string in Object
'
' @param     {string}    strQueryStringParameters  String
'
' @return    {strQueryStringParameters}
' @throws    invalidQueryStringParameters
Public Function parseQueryStringParameters( _
       ByVal strQueryStringParameters As String _
)

    Dim response As New dictionary
    Dim queryKeys As Collection


    rowsQuery = Split(strQueryStringParameters, "&")
    
    For i = LBound(rowsQuery) To UBound(rowsQuery)
    
        Set rowQueryStringParameters = splitKeyValue(rowsQuery(i))
        Set queryKeys = getKeys(rowQueryStringParameters.Item("key"))
        
        principalKey = queryKeys.Item(1): queryKeys.Remove 1
        If (queryKeys.count = 0) Then
            If (response.Exists(principalKey)) Then
                If (IsObject(rowQueryStringParameters.Item("value"))) Then
                    Set response.Item(principalKey) = rowQueryStringParameters.Item("value")
                Else
                    response.Item(principalKey) = rowQueryStringParameters.Item("value")
                End If
            Else
                response.add principalKey, rowQueryStringParameters.Item("value")
            End If
        Else
            
            If (response.Exists(principalKey)) Then
                If (TypeName(response.Item(principalKey)) <> classNameDictionary) Then
                    Set response.Item(principalKey) = New dictionary
                End If
            Else
                response.add principalKey, New dictionary
            End If
            
            Set lastSubKey = response.Item(principalKey)
            
            keysConcat = Array()
            For Each subKey In queryKeys

                If (subKey = "") Then
                    subKeyName = sequentialKey
                Else
                    subKeyName = subKey & ""
                End If

                keysConcat = arrayMerge(keysConcat, Array(subKeyName))
                
                If (TypeName(lastSubKey.Item(subKeyName)) = classNameDictionary) Then
                    Set lastSubKey = lastSubKey.Item(subKeyName)
                Else
                    Set lastSubKey.Item(subKeyName) = New dictionary: Set lastSubKey = lastSubKey.Item(subKeyName)
                End If
            Next
            Set response.Item(principalKey) = assingDictionary(response.Item(principalKey), keysConcat, rowQueryStringParameters.Item("value"))

        End If
        
    
    Next

    Set parse = response

End Function

''
' Split a row in two params, key and value
'
' @param {string}    rowValue    Value to Split
' @return {Collection}
'
' @throws emptyField
Private Function splitKeyValue(ByVal rowValue As String) As dictionary
    
    Dim response As New dictionary
    
    splited = Split(rowValue, "=")
    
    If (Not isIterable(splited)) Then
        throw QueryStringParametersThrowsCode.emptyField
    Else
        
        QueryValue = ""
        If (UBound(splited) > 1) Then: QueryValue = splited(1)
    
        response.add "key", splited(0)
        response.add "value", splited(1)
    
    End If

    Set splitKeyValue = response
    
End Function

''
' Get keys from string query parameters
'
' @param    {string}    stringToSearch  Query String to get Keys
'
' @return   {Collection}
Private Function getKeys( _
         ByVal stringToSearch As String _
) As Collection

    Dim response As New Collection
    Set thisRegex = getRegex()

    stringToSearch = decodeUrl(stringToSearch)
    stringToSearch = Replace(stringToSearch, "[]", "[" & SEQUENTIAL_KEY & "]")
    
    response.add Mid(stringToSearch, 1, IIf(InStr(1, stringToSearch, "[") = 0, Len(stringToSearch), InStr(1, stringToSearch, "[") - 1))

    expression = "^.*"
    replaceExpression = ""
    For i = 1 To Len(stringToSearch) - Len(Replace$(stringToSearch, "[", ""))
        expression = expression & "\[([^\]]*)\]"
        replaceExpression = replaceExpression & "$" & i & "|{%delimiter%}|"
    Next
    
    If (Len(replaceExpression) > 1) Then: replaceExpression = Mid(replaceExpression, 1, Len(replaceExpression) - Len("|{%delimiter%}|"))
    If (Len(expression) > 3) Then
        expression = expression & "$"
        With thisRegex
            .Global = False
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = expression
            valueInField = Split(.Replace(stringToSearch, replaceExpression), "|{%delimiter%}|")
        End With

        If (isIterable(valueInField)) Then
    
            For i = LBound(valueInField) To UBound(valueInField)
                response.add IIf(valueInField(i) = SEQUENTIAL_KEY, "", valueInField(i))
            Next
            
        End If
    End If

    Set getKeys = response

End Function

''
' Get a sequential key in query string parsed
'
' @param    {Double}    Index       Index taht you wish get
' @param    {String()}  arrayKeys   Array keys to level that you wish get index
'
' @return   {String}    return "-1" if not found
Public Function getSequentialKeyByIndex(ByVal index As Double, Optional ByVal arrayKeys = "") As String
    
    Dim count As Double: count = -1
    Dim dictionary As dictionary
    
    If (isIterable(arrayKeys)) Then
        Dim dicResponse As New dictionary
        dicResponse.add "value", getLineageValueDictionary(parse, arrayKeys)
        
        If (TypeName(dicResponse.Item("value")) <> classNameDictionary) Then
            Set dictionary = New dictionary
        Else
            Set dictionary = dicResponse.Item("value")
        End If
    Else
        Set dictionary = parse
    End If
    
    getSequentialKeyByIndex = -1
    
    dKeys = dictionary.keys()
    If (UBound(dKeys) <> -1) Then
        For i = LBound(dKeys) To UBound(dKeys)
            If (isSequentialKey(dKeys(i))) Then
                count = count + 1
            End If
            
            If (count = index) Then
                getSequentialKeyByIndex = dKeys(i)
                Exit For
            End If
        Next
    End If
    
End Function

''
' Get a object Regex
'
' @return {Regex}
Private Function getRegex()
    If Regex Is Nothing Then
        Set Regex = CreateObject("VBScript.RegExp")
    End If
    
    Set getRegex = Regex
End Function

''
' Encode string to URI string
'
' @param     {string}    varText     Text to encode
'
' @return    {string}
Public Function encodeUrl(ByVal varText As String)

    If Htmlfile Is Nothing Then
        Set Htmlfile = CreateObject("htmlfile")
    End If
    With Htmlfile.parentWindow
        .execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
        encodeUrl = .encode(varText)
    End With
    
End Function

''
' Decode URI string
'
' @param     {string}    varText     Text to decode
'
' @return    {string}
Public Function decodeUrl(ByVal varText As String)

    If Htmlfile Is Nothing Then
        Set Htmlfile = CreateObject("htmlfile")
    End If
    With Htmlfile.parentWindow
        .execScript "function decode(s) {return decodeURIComponent(s)}", "jscript"
        decodeUrl = .decode(varText)
    End With
    
End Function

''
' Escape values to regex
'
' @param    {string}    value       Value to Escape
' @param    {string}    ?delimiter  delimiter used
'
' @return   {string}
Private Function escapeRegex(ByVal value As String, Optional ByVal demiliter As String)

    If Htmlfile Is Nothing Then
        Set Htmlfile = CreateObject("htmlfile")
    End If
    With Htmlfile.parentWindow
        .execScript "function escapeRegex(str, delimiter) {return (str + '').replace(new RegExp('[.\\\\+*?\\[\\^\\]\\$(){}=!<>|:\\' + (delimiter || '') + '-]', 'g'), '\\$&');}", "jscript"
        escapeRegex = .escapeRegex(value)
    End With

End Function

''
' Search a vaue in a Array
'
' @param    {*}         needle      Searched
' @param    {array}     haystack    Array
' @return   {Double}    If not found return -1
Private Function arraySearch(ByVal needle As Variant, ByVal haystack) As Double
    
    If (Not (isIterable(haystack))) Then
        arraySearch = -1
    End If

    For i = LBound(haystack) To UBound(haystack)
        If (needle = haystack(i)) Then: Exit For
    Next

    arraySearch = IIf(i > UBound(haystack), -1, i)
    End Function

    ''
    ' Merge a array
    '
    ' @param    {array} ...arraysToMerge    array to merge
    ' @return   {array}
Private Function arrayMerge(ParamArray arraysToMerge() As Variant) As Variant()

    Dim response() As Variant

    totalIndex = -1
    setValues = False
    
row__init:

    realIndex = 0
    For i = LBound(arraysToMerge) To UBound(arraysToMerge)
        
        arrayToMerge = arraysToMerge(i)
        
        If (isIterable(arrayToMerge)) Then
                For ii = LBound(arrayToMerge) To UBound(arrayToMerge)
                
                    If (setValues) Then
                        response(realIndex) = arrayToMerge(ii)
                    Else
                        totalIndex = totalIndex + 1
                    End If
                        
                    realIndex = realIndex + 1
                    
                Next
        End If
    Next
    
    If (totalIndex = -1) Then
        response = Array()
    ElseIf (Not setValues) Then
        ReDim response(0 To totalIndex)
        setValues = True
        GoTo row__init
    End If
    
    arrayMerge = response
    
End Function

''
' Slice a Array
'
' @param    {Variant()}     arraySlice  Array
' @param    {Double}        startSlice  Init to slice the array
' @param    {Double}        endSlice?   End of Slice
'
' @return   {Variant()}
Private Function arraySlice(ByVal arrayToSlice, ByVal startSlice As Double, Optional ByVal endSlice As Double = -1)

    Dim response()

    If (Not (isIterable(arrayToSlice))) Then
        arraySlice = Array()
    End If

    startIndex = LBound(arrayToSlice)
    endIndex = UBound(arrayToSlice)

    endSlice = IIf(endSlice = -1, endIndex, endSlice)
    
    ReDim response(0 To endSlice - startSlice)
    newIndex = 0
    For i = startIndex To endIndex
        If (i >= startSlice And i <= endSlice) Then
            response(newIndex) = arrayToSlice(i)
            newIndex = newIndex + 1
        End If
    Next

    arraySlice = response

End Function

''
' Assign a new value in a keys from a Dictionary
'
' @param    {Dictionary}    dictionaryToAssing  Dictionary
' @param    {String()}      keysToAssing        Keys to assing a new value
' @param    {Variant}       value               New value
'
' @return   {Dictionary}
Private Function assingDictionary( _
        ByVal dictionaryToAssing As dictionary, _
        ByVal keysToAssing, _
        ByVal value, _
        Optional ByVal after = "" _
) As dictionary
    
    Dim alreadyAdded As Boolean
    after = IIf(after = "", addAsLastValue, after)
    
    If (0 = UBound(keysToAssing)) Then
    
        Dim newDictionaryToAssing As New dictionary
        
        If (after = addAsFirstValue) Then
            newDictionaryToAssing.add keysToAssing(0), value
            alreadyAdded = True
        End If
        
        For Each key In dictionaryToAssing.keys
        
            newDictionaryToAssing.add key, dictionaryToAssing.Item(key)
            
            If (key = after) Then
                If (IsObject(value)) Then
                    newDictionaryToAssing.Item(keysToAssing(0)) = value
                Else
                    newDictionaryToAssing.Item(keysToAssing(0)) = value
                End If
                alreadyAdded = True
            End If
        Next
        
        If (after = addAsLastValue Or Not alreadyAdded) Then
            If (IsObject(value)) Then
                newDictionaryToAssing.Item(keysToAssing(0)) = value
            Else
                newDictionaryToAssing.Item(keysToAssing(0)) = value
            End If
        End If
        
        Set dictionaryToAssing = newDictionaryToAssing
    Else
        
        If (Not TypeName(dictionaryToAssing.Item(keysToAssing(0))) = classNameDictionary) Then
            Set dictionaryToAssing.Item(keysToAssing(0)) = New dictionary
        End If
        
        Set subDic = dictionaryToAssing.Item(keysToAssing(0))
        
        Set dictionaryToAssing.Item(keysToAssing(0)) = assingDictionary(subDic, arraySlice(keysToAssing, 1), value, after)
        
    End If

    Set assingDictionary = dictionaryToAssing

End Function

''
' Assign a new value in a keys from a Dictionary
'
' @param    {Dictionary}    dictionaryToAssing  Dictionary
' @param    {String()}      keysToAssing        Keys to assing a new value
'
' @return   {Variant|Error} return error if not Found
Private Function getLineageValueDictionary( _
        ByVal dictionary As dictionary, _
        ByVal keysToAssing _
)
    
    Dim dic As New dictionary
    
    If (0 = UBound(keysToAssing)) Then
        If (dictionary.Exists(keysToAssing(0))) Then
            If (IsObject(dictionary.Item(keysToAssing(0)))) Then
                Set getLineageValueDictionary = dictionary.Item(keysToAssing(0))
            Else
                Let getLineageValueDictionary = dictionary.Item(keysToAssing(0))
            End If
        Else
            getLineageValueDictionary = CVErr(2015)
        End If
    ElseIf (0 < UBound(keysToAssing)) Then
        newKeys = arraySlice(keysToAssing, 1)
        
        If (dictionary.Exists(keysToAssing(0))) Then
            Set subDictionary = IIf(TypeName(dictionary.Item(keysToAssing(0))) <> classNameDictionary, New dictionary, dictionary.Item(keysToAssing(0)))
        Else
            Set subDictionary = New dictionary
        End If
        Set response = New dictionary
        
        response.add "value", getLineageValueDictionary(subDictionary, newKeys)
        
        If (IsObject(response.Item("value"))) Then
            Set getLineageValueDictionary = response.Item("value")
        Else
            getLineageValueDictionary = response.Item("value")
        End If
    End If
    
End Function

Private Function removeLineageValueDictionary( _
        ByVal dictionary As dictionary, _
        ByVal keysToAssing _
) As dictionary
    
    Dim dic As New dictionary
    
    If (isIterable(keysToAssing)) Then
    
        If (0 = UBound(keysToAssing)) Then
            If (dictionary.Exists(keysToAssing(0))) Then
                dictionary.Remove keysToAssing(0)
            End If
        ElseIf (0 < UBound(keysToAssing)) Then
        
            newKeys = arraySlice(keysToAssing, 1)
            
            If (dictionary.Exists(keysToAssing(0))) Then
                Set subDictionary = IIf(TypeName(dictionary.Item(keysToAssing(0))) <> classNameDictionary, New dictionary, dictionary.Item(keysToAssing(0)))
            Else
                Set subDictionary = New dictionary
            End If
            
            Set dictionary.Item(keysToAssing(0)) = removeLineageValueDictionary(subDictionary, newKeys)
            
        End If
    End If
    
    Set removeLineageValueDictionary = dictionary
    
End Function

''
' Get parsed query string parameters in string
'
' @return {string}
' @throws unknowError
Public Function toString() As String
    
    Dim response()
    
    For Each key In parse.keys
        
        If (TypeName(parse(key)) = classNameDictionary) Then
            Set dicKeys = getSubValue(parse(key), key)
            For Each entireKey In dicKeys.keys
                response = arrayMerge(response, Array(generateKeyValue(entireKey, dicKeys(entireKey))))
            Next
        ElseIf (IsObject(parse(key))) Then
            throw unknowError, "Maybe property ""parse"" has changed manually, has a value in ""parsed"" that your type name is """ & TypeName(parse(key)) & """"
        Else
            response = arrayMerge(response, Array(generateKeyValue(key, parse(key))))
        End If
    Next
    
    With getRegex()
        .Pattern = escapeRegex(encodeUrl("[")) & "(" & getRegexSequentialKey() & ")" & escapeRegex(encodeUrl("]"))
        .Global = True
        entireString = Join(response, "&")
        toString = .Replace(entireString, encodeUrl("[]"))
    End With
    
End Function
''
' get sub keys with value from a dictionary
'
' @param    {dictionary}    dic         Dictionary
' @param    {string}        parentKeys  parent keys from sub keys
'
' @return   {dictionary}
' @throws   unknowError
Private Function getSubValue(ByVal dic As dictionary, Optional ByVal parentKeys As String = "") As dictionary
    
    Dim response As New dictionary
    
    For Each key In dic.keys
        
        exportKey = IIf(isSequentialKey(key), "", key)
        
        If (TypeName(dic(key)) = classNameDictionary) Then
            Set subKeys = getSubValue(dic(key), parentKeys & "[" & key & "]")
            For Each subKey In subKeys.keys
                response.Item(subKey) = subKeys(subKey)
            Next
        ElseIf (IsObject(dic(key))) Then
            throw unknowError, "Maybe property ""parse"" has changed manually, has a value in ""parsed"" that your type name is """ & TypeName(parse(key)) & """"
        Else
            response.Item(parentKeys & "[" & key & "]") = dic(key)
        End If
    Next
    
    Set getSubValue = response
    
End Function
''
' Get key and value to form Query Params
'
' @param    {string}    key     Parameter key
' @param    {string}    value   Parameter value
'
' @return   {array}
Private Function generateKeyValue(ByVal key As String, ByVal value As String) As String
    generateKeyValue = encodeUrl(key) & "=" & encodeUrl(value)
End Function
''
' Check if variable is iterable
'
' @param    {array}
'
' @return   {boolean}
Private Function isIterable(ByVal arr) As Boolean
On Error GoTo catch
    isIterable = UBound(arr) > -1
    Exit Function
catch:
    isIterable = False
End Function

''
' Add a new value in params
'
' @param    {string}     value      Value to assign a key
' @param    {string()}   arrayKeys  Lineage keys to assing
'
' @throws   invalidValueToAssing
Public Function add(ByVal value, arrayKeys, Optional ByVal after As String = "")

    If (isIterable(arrayKeys)) Then
        For i = LBound(arrayKeys) To UBound(arrayKeys)
            If (isValidKey(arrayKeys(i))) Then
            
                arrayKeys(i) = arrayKeys(i) & ""
                
                If (IsNumeric(arrayKeys(i))) Then
                    sequentialKeyIndex = getSequentialKeyByIndex(arrayKeys(i))
                    arrayKeys(i) = IIf(sequentialKeyIndex = -1, arrayKeys(i), sequentialKeyIndex)
                ElseIf (arrayKeys(i) = "") Then
                    arrayKeys(i) = sequentialKey
                End If
                
            Else
                throw invalidValueToAssing, "arrayKeys must be string or numeric, actual type: " & TypeName(arrayKeys(i))
            End If
        Next
    Else
        throw invalidValueToAssing, "specify a array keys!"
    End If
    
    If (Not IsError(getLineageValueDictionary(parse, arrayKeys))) Then
        throw valueAlreadyExists
    End If
    
    If (isValidValue(value)) Then
        Set parse = assingDictionary(parse, arrayKeys, value, after)
    Else
        throw invalidValueToAssing, "var type must be string or numeric, actual type: " & TypeName(value)
    End If
End Function

''
' Set a value in new field, add if not exists
'
' @param    {*}     value       Value to update
' @param    {*}     arrayKeys   Associetive keys to update value
'
' @throws   invalidValueToAssing
Public Function update(ByVal value, arrayKeys, Optional ByVal after = "")
    
    If (isIterable(arrayKeys)) Then
        For i = LBound(arrayKeys) To UBound(arrayKeys)
            If (isValidKey(arrayKeys(i))) Then
                
                arrayKeys(i) = arrayKeys(i) & ""
                
                If (arrayKeys(i) = "") Then
                    arrayKeys(i) = sequentialKey
                End If
                
            Else
                throw invalidValueToAssing, "keys must be string or numeric, actual type: " & TypeName(arrayKeys(i))
            End If
        Next
    Else
        throw invalidValueToAssing, "specify a array keys!"
    End If
    
    If (isValidValue(value)) Then
        Set parse = assingDictionary(parse, arrayKeys, value, after)
    Else
        throw invalidValueToAssing, "var type must be string or numeric, actual type: " & TypeName(value)
    End If
    
End Function

''
' Get a value through a array keys associative
'
' @param    {*}  keys            Array keys associative
' @param    {*}  ?replacement    If not found, return this
'
' @return   {*|Error(2015)}
Public Function getValue(ByVal keys, Optional ByVal replacement = "&-1n4liD_%=r3PL4Ã§3m3nt-&")

    Dim response As New dictionary

    If (isIterable(keys)) Then
        For i = 0 To UBound(keys)
            If (isValidKey(keys(i))) Then
                keys(i) = keys(i) & ""
            Else
                throw invalidValueToAssing, "keys must be string or numeric, actual type: " & TypeName(keys(i))
            End If
        Next
    Else
        throw invalidValueToAssing, "specify a array keys!"
    End If

    response.add "preValue", getLineageValueDictionary(parse, keys)

    
    If (IsObject(replacement)) Then
        haveReplacement = True
    ElseIf (replacement <> "&-1n4liD_%=r3PL4Ã§3m3nt-&") Then
        haveReplacement = False
    Else
        haveReplacement = True
    End If
    
    
    If (IsError(response.Item("preValue")) And haveReplacement) Then
        response.add "value", replacement
    Else
        response.add "value", response.Item("preValue")
    End If
    
    If (IsObject(response.Item("value"))) Then
        Set getValue = response.Item("value")
    Else
        getValue = response.Item("value")
    End If
    
End Function
''
' Delete a value through a array keys associative
'
' @param    {*}  keys            Array keys associative
Public Function delete(ByVal arrayKeys)

    If (isIterable(arrayKeys)) Then
        For i = 0 To UBound(arrayKeys)
            If (isValidKey(arrayKeys(i))) Then
                arrayKeys(i) = arrayKeys(i) & ""
            Else
                throw invalidValueToAssing, "keys must be string or numeric, actual type: " & TypeName(arrayKeys(i))
            End If
        Next
    Else
        throw invalidValueToAssing, "specify a array keys!"
    End If
    
    Set parse = removeLineageValueDictionary(parse, arrayKeys)
    
End Function
''
' Check if key is valid
'
' @param    {*}     key     Key
'
' @return   {Boolean}
Private Function isValidKey(ByVal key) As Boolean
    isValidKey = "STRING|INTEGER|DOUBLE|DATE" Like ("*" & UCase(TypeName(key)) & "*")
End Function

''
' Check if value is valid
'
' @param    {*}     value   Value
'
' @return   {Boolean}
Private Function isValidValue(ByVal value) As Boolean
    isValidValue = "STRING|INTEGER|DOUBLE|DATE" Like ("*" & UCase(TypeName(value)) & "*")
End Function

''
' Get parse
'
' @return dictionary
Public Function getParsed()
    Set getParsed = parse
End Function

''
' Get a expression use to add a sequential key
'
' @return   {string}
Private Function sequentialKey() As String
    sequentialKey = SEQUENTIAL_KEY & "__" & Replace(Replace("" & (Now * 1) & Int(1 + Rnd * (9999 - 1 + 1)), ",", ""), ".", "") & "__"
End Function

''
' Check if expression is sequencial key
'
' @param    {string}    expression  Expression to check
'
' @return   {boolean}
Private Function isSequentialKey(ByVal expression As String) As Boolean
    
    With getRegex()
        .Pattern = "^" & getRegexSequentialKey() & "$"
        isSequentialKey = .test(expression)
    End With
    
End Function

''
' Get Pattern used to check sequentialKey
'
' @return   {String}
Private Function getRegexSequentialKey() As String
    getRegexSequentialKey = escapeRegex(SEQUENTIAL_KEY) & "__\d+__"
End Function

''
' Initialize Class
Private Sub Class_Initialize()
    addAsFirstValue = "[___first_value___]"
    addAsLastValue = "[___last_value___]"
End Sub
