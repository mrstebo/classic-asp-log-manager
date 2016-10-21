<%
' So we don't end up with null reference-type exceptions :)
Class NullLogger
	Private pName
	Private pLogLevel

	Public Property Let LogLevel(value)
		pLogLevel = value
	End Property

	Private Sub Class_Initialize
		pLogLevel = LogLevels.FATAL
	End Sub

	Public Default Function construct(name)
		pName = name
		Set construct = Me
	End Function

	Public Sub Log(logObject)
	End Sub
End Class
%>
