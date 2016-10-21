<%
Class SimpleLogObject
  Public LogLevel
  Public Message

  Public Function ToJSON()
    Dim result : result = ""

    result = result & "{"
    result = result & """level"": "&LogLevel&","
    result = result & """message"": """&JsonUtils.MakeJsonSafe(Message)&""""
    result = result & "}"

    ToJson = result
  End Function
End Class
%>
