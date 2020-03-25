Sub ListaNomes()
    Dim i As Long
    Dim intCount As Long
    Dim NameArray() As String
    'count number of named ranges
    intCount = ThisWorkbook.Names.Count
    ReDim NameArray(intCount)  'for dynamic number of named ranges
    'assign all names to each array element
    For i = 1 To intCount
      NameArray(i) = ThisWorkbook.Names(i).Name 'remove .Name if you want address range only
    Next
   'check array elements
   For Each ele In NameArray
     Debug.Print ele
   Next
  'manipulate your array here
  End Sub