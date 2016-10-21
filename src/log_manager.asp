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

  Private Function FindOrCreateLogger(logger)
    Dim loggerName : loggerName = logger.Name

    If pLoggers.Exists(loggerName) Then
      If pLoggers(loggerName) Is Nothing Then
        pLoggers.Remove(loggerName)
      End If
    End If

    If Not pLoggers.Exists(loggerName) Then
      Call pLoggers.Add(loggerName, logger)
    End If

    Set AppendLogger = pLoggers(loggerName)
  End Function
End Class

Dim LogManager : Set LogManager = New LogManager_
%>
