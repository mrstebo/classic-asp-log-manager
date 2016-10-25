<%
Class LogBuilder
  Private pLog

  Private Sub Class_Initialize
    Set pLog = New LogObject
  End Sub

  Public Function Trace(message)
    pLog.LogLevel = 1
    pLog.Message = message
    Set Trace = Me
  End Function

  Public Function Debug(message)
    pLog.LogLevel = 2
    pLog.Message = message
    Set Debug = Me
  End Function

  Public Function Info(message)
    pLog.LogLevel = 3
    pLog.Message = message
    Set Info = Me
  End Function

  Public Function Warning(message)
    pLog.LogLevel = 4
    pLog.Message = message
    Set Warning = Me
  End Function

  Public Function Fatal(message)
    pLog.LogLevel = 5
    pLog.Message = message
    Set Fatal = Me
  End Function

  Public Function Meta(key, value)
    Call pLog.Meta.Add(key, value)
    Set Meta = Me
  End Function

  Public Function Tag(tagName)
    pLog.Tags.AddItem tagName
    Set Tag = Me
  End Function

  Public Function Build()
    Set Build = pLog
  End Function

  ' Classed as "extension methods"

  Public Function ExceptionMeta(ex)
    On Error Resume Next
    Call Meta("exception[Number]", ex.Number)
    Call Meta("exception[Source]", ex.Source)
    Call Meta("exception[Message]", ex.Description)
    Set ExceptionMeta = Me
  End Function

  Public Function RequestMeta()
    On Error Resume Next
    Call Meta("request_data[Url]", UrlUtils.GetCurrentUrl())
    Call Meta("session[SessionID]", Session.SessionID)
    Call Meta("server_variables[ALL_RAW]", Request.ServerVariables("ALL_RAW"))
    Call Meta("server_variables[CONTENT_LENGTH]", Request.ServerVariables("CONTENT_LENGTH"))
    Call Meta("server_variables[CONTENT_TYPE]", Request.ServerVariables("CONTENT_TYPE"))
    Call Meta("server_variables[GATEWAY_INTERFACE]", Request.ServerVariables("GATEWAY_INTERFACE"))
    Call Meta("server_variables[HTTP_ACCEPT]", Request.ServerVariables("HTTP_ACCEPT"))
    Call Meta("server_variables[HTTP_ACCEPT_LANGUAGE]", Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"))
    Call Meta("server_variables[HTTP_COOKIE]", Request.ServerVariables("HTTP_COOKIE"))
    Call Meta("server_variables[HTTP_REFERER]", Request.ServerVariables("HTTP_REFERER"))
    Call Meta("server_variables[HTTP_USER_AGENT]", Request.ServerVariables("HTTP_USER_AGENT"))
    Call Meta("server_variables[HTTPS]", Request.ServerVariables("HTTPS"))
    Call Meta("server_variables[INSTANCE_ID]", Request.ServerVariables("INSTANCE_ID"))
    Call Meta("server_variables[INSTANCE_META_PATH]", Request.ServerVariables("INSTANCE_META_PATH"))
    Call Meta("server_variables[LOCAL_ADDR]", Request.ServerVariables("LOCAL_ADDR"))
    Call Meta("server_variables[QUERY_STRING]", Request.ServerVariables("QUERY_STRING"))
    Call Meta("server_variables[REMOTE_ADDR]", Request.ServerVariables("REMOTE_ADDR"))
    Call Meta("server_variables[REMOTE_HOST]", Request.ServerVariables("REMOTE_HOST"))
    Call Meta("server_variables[REMOTE_USER]", Request.ServerVariables("REMOTE_USER"))
    Call Meta("server_variables[REQUEST_METHOD]", Request.ServerVariables("REQUEST_METHOD"))
    Call Meta("server_variables[SCRIPT_NAME]", Request.ServerVariables("SCRIPT_NAME"))
    Call Meta("server_variables[SERVER_NAME]", Request.ServerVariables("SERVER_NAME"))
    Call Meta("server_variables[SERVER_PORT]", Request.ServerVariables("SERVER_PORT"))
    Call Meta("server_variables[SERVER_PORT_SECURE]", Request.ServerVariables("SERVER_PORT_SECURE"))
    Call Meta("server_variables[SERVER_PROTOCOL]", Request.ServerVariables("SERVER_PROTOCOL"))
    Call Meta("server_variables[SERVER_SOFTWARE]", Request.ServerVariables("SERVER_SOFTWARE"))
    Call Meta("server_variables[URL]", Request.ServerVariables("URL"))
    Set RequestMeta = Me
  End Function
End Class
%>
