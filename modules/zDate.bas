Attribute VB_Name = "zDate"
Function GetDateStr() As String
  GetDateStr = FormatDateTime(Now, vbLongDate)
End Function

