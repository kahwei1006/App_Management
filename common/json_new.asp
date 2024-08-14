<%
Class aspJSON

    ' Object to hold JSON data
    Private jsonData
    
    ' Constructor
    Private Sub Class_Initialize()
        Set jsonData = CreateObject("Scripting.Dictionary")
    End Sub
    
    ' Destructor
    Private Sub Class_Terminate()
        Set jsonData = Nothing
    End Sub
    
    ' Method to load JSON string and parse it
    Public Sub loadJSON(jsonString)
        Dim dict, i, arr
        jsonString = Replace(jsonString, vbCrLf, "")
        jsonString = Replace(jsonString, vbTab, "")
        jsonString = Replace(jsonString, "},", "}|")
        arr = Split(Mid(jsonString, 2, Len(jsonString)-2), "|")
        For i = 0 To UBound(arr)
            Set dict = ParseJSONObject(Trim(arr(i)))
            jsonData.Add i, dict
        Next
    End Sub
    
    ' Method to parse individual JSON objects
    Private Function ParseJSONObject(jsonObject)
        Dim dict, pairs, pair, key, value, i
        Set dict = CreateObject("Scripting.Dictionary")
        jsonObject = Replace(jsonObject, "{", "")
        jsonObject = Replace(jsonObject, "}", "")
        pairs = Split(jsonObject, ",")
        For i = 0 To UBound(pairs)
            pair = Split(pairs(i), ":")
            key = Replace(Trim(pair(0)), """", "")
            value = Replace(Trim(pair(1)), """", "")
            dict.Add key, value
        Next
        Set ParseJSONObject = dict
    End Function
    
    ' Method to get the jsonData object
    Public Function getData()
        Set getData = jsonData
    End Function
    
    ' Method to get the count of items
    Public Function count()
        count = jsonData.Count
    End Function
    
    ' Method to get a specific item
    Public Function item(index)
        Set item = jsonData.Item(index)
    End Function

End Class
%>