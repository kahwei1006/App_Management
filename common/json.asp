<%
Class aspJSON

    ' Method to load and parse JSON string into a dictionary
    Public Function loadJSON(strJSON)
        Dim dict, pairs, pair, keyValue, i
        Set dict = Server.CreateObject("Scripting.Dictionary")

        ' Clean up the JSON string
        strJSON = Replace(strJSON, "{", "")
        strJSON = Replace(strJSON, "}", "")
        strJSON = Replace(strJSON, Chr(13), "") ' Remove carriage returns
        strJSON = Replace(strJSON, Chr(10), "") ' Remove newlines
        strJSON = Replace(strJSON, " ", "")     ' Remove spaces

        pairs = Split(strJSON, ",")

        For i = 0 To UBound(pairs)
            pair = pairs(i)
            keyValue = Split(pair, ":")
            If UBound(keyValue) = 1 Then
                dict(Trim(Replace(keyValue(0), """", ""))) = Trim(Replace(keyValue(1), """", ""))
            End If
        Next

        Set loadJSON = dict
    End Function

End Class
%>
