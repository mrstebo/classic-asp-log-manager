<%
Class GuidGenerator_
  Private pNextSeed

  Private Sub Class_Initialize
    pNextSeed = Timer
  End Sub

  Public Function Create()
    Create = ""
    Create = Create & CreateFixedLength(8) & "-"
    Create = Create & CreateFixedLength(4) & "-"
    Create = Create & CreateFixedLength(4) & "-"
    Create = Create & CreateFixedLength(4) & "-"
    Create = Create & CreateFixedLength(12)
  End Function

  Private Function CreateFixedLength(length)
    Const valid = "0123456789abcdef"
    Dim guid
    Dim i
    For i = 1 To length
      guid = guid & Mid(valid, Int(GetRandomNumber() * Len(valid)) + 1, 1)
    Next
    CreateFixedLength = guid
  End Function

  ' Need to generate another seed after getting a random number so we are
  ' less likely to end up with the same combination
  Private Function GetRandomNumber()
    Randomize pNextSeed
    GetRandomNumber = Rnd(1)
    pNextSeed = Rnd(1) * 100000
  End Function
End Class

Dim GuidGenerator : Set GuidGenerator = New GuidGenerator_
%>
