<%
Class ArrayList
	Private pArrayList

	Public Property Get Count
		Count = pArrayList.Count
	End Property

	Private Sub Class_Initialize
		Set pArrayList = CreateObject("System.Collections.ArrayList")
	End Sub

	Private Sub Class_Terminate
		Set pArrayList = Nothing
	End Sub

	Public Function GetAll()
		Set GetAll = pArrayList
	End Function

	Public Function GetByIndex(index)
		If VarType(pArrayList(index)) = vbObject Then
			Set GetByIndex = pArrayList(index)
		Else
			GetByIndex = pArrayList(index)
		End If
	End Function

	Public Sub AddItem(obj)
		pArrayList.add obj
	End Sub

	Public Function ToArray()
		Dim results()
		If pArrayList.Count > 0 Then
			ReDim results(pArrayList.Count - 1)

			Dim i
			For i = 0 To UBound(results)
				If VarType(pArrayList(i)) = vbObject Then
					Set results(i) = pArrayList(i)
				Else
					results(i) = pArrayList(i)
				End If
			Next
		End If

		ToArray = results
	End Function
End Class
%>
