Attribute VB_Name = "modDate"
Function GetDateStr() As String
  GetDateStr = FormatDateTime(Now, vbLongDate)
End Function

