<%
Class LogManager_
  Private pRequestId
  Private pLoggers

  Public Property Get RequestID
    RequestID = pRequestId
  End Property

  Private Sub Class_Initialize
    pRequestId = GuidGenerator.Create()
    Set pLoggers = CreateObject("Scripting.Dictionary")
  End Sub

  Private Sub Class_Terminate
    For Each key in pLoggers.Keys
      pLoggers(key) = Nothing
    Next
    pLoggers.RemoveAll
    Set pLoggers = Nothing
  End Sub

  Public Function GetLogger(loggerName)
    Set GetLogger = FindOrCreateLogger(loggerName, (New NullLogger)(loggerName))
  End Function

  Public Function GetFileLogger(loggerName, filename)
    Set GetFileLogger = FindOrCreateLogger(loggerName, (New FileLogger)(loggerName, filename))
  End Function

  Private Function FindOrCreateLogger(loggerName, loggerFactory)
    If pLoggers.Exists(loggerName) Then
      If pLoggers(loggerName) Is Nothing Then
        pLoggers.Remove(loggerName)
      End If
    End If

    If Not pLoggers.Exists(loggerName) Then
      Call pLoggers.Add(loggerName, loggerFactory)
    End If

    Set FindOrCreateLogger = pLoggers(loggerName)
  End Function
End Class

Dim LogManager : Set LogManager = New LogManager_
%>
