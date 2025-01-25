Function GetObjectRobust(progID)
  On Error GoTo ErrHandler
  Set obj = GetObject(progID)
  Exit Function
ErrHandler:
  Err.Clear
  On Error Resume Next
  Set obj = CreateObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    MsgBox "Error creating or getting object: " & progID & ". Error number: " & Err.Number, vbCritical
    Set obj = Nothing ' Explicitly set to Nothing to avoid ambiguity
  End If
  On Error GoTo 0
  Set GetObjectRobust = obj
End Function

Sub Main
  Dim myObj As Object
  Set myObj = GetObjectRobust("Some.Invalid.ProgID")
  If Not myObj Is Nothing Then
    Debug.Print myObj.SomeProperty ' This is less likely to throw an error because the function is improved 
  Else
    Debug.Print "Failed to create or get object."
  End If
End Sub