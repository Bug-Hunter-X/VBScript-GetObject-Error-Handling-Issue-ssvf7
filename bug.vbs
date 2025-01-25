Function GetObject(progID) 
  On Error Resume Next 
  Set obj = GetObject(progID) 
  If Err.Number <> 0 Then 
    Err.Clear 
    Set obj = CreateObject(progID) 
  End If 
  Set GetObject = obj
End Function

'This function is designed to retrieve an object, first by trying to get it from the existing objects and then creating it if it does not exist.
'However, the error handling is imperfect. If CreateObject fails for any reason (e.g., the progID is incorrect or the necessary component is not installed), it will set obj to Nothing and then return Nothing, without indicating any failure. This can lead to unexpected behavior down the line because calling methods on a Nothing object will result in runtime errors.

Sub Main
  Dim myObj As Object
  Set myObj = GetObject("Some.Invalid.ProgID")
  If Not myObj Is Nothing Then
    'This line will throw an error if GetObject failed and returned Nothing
    Debug.Print myObj.SomeProperty
  End If
End Sub