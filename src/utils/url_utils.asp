<%
Class UrlUtils_
  Public Function GetCurrentUrl()
    Dim result : result = ""
    Dim query : query = Request.ServerVariables("QUERY_STRING")

    result = result & "http"

    If Request.ServerVariables("HTTPS") = "on" Then
      result = result & "s"
    End If

    result = result & "://"
    result = result & Request.ServerVariables("HTTP_HOST")
    result = result & Request.ServerVariables("PATH_INFO")

    If Len(query) > 0 Then
      result = result & "?" & query
    End If

    GetCurrentUrl = result
  End Function
End Class

Dim UrlUtils : Set UrlUtils = New UrlUtils_
%>
