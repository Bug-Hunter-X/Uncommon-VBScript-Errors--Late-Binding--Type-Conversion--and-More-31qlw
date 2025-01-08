Function CheckObjectSupport(obj, methodName)
  On Error Resume Next
  Dim supported
  supported = obj.methodName
  On Error GoTo 0
  If Err.Number = 0 Then
    CheckObjectSupport = True
  Else
    CheckObjectSupport = False
    Err.Clear
  End If
End Function

Function SafeStringToInt(str)
  Dim num
  On Error Resume Next
  num = CLng(str)
  On Error GoTo 0
  If Err.Number = 0 Then
    SafeStringToInt = num
  Else
    SafeStringToInt = 0  ' Or handle the error appropriately
    Err.Clear
  End If
End Function

'Example usage
Dim myObject: Set myObject = CreateObject("Scripting.Dictionary")
If CheckObjectSupport(myObject, "Add") Then
    MsgBox "Add method supported"
Else
    MsgBox "Add method NOT supported"
End If

Dim strNum: strNum = "10abc"
Dim num: num = SafeStringToInt(strNum)
MsgBox "Converted number: " & num