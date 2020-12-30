Option Explicit

Private c As Collection
Private dataSizeIndex As Long
Private firstDataPosition As Long

Private Sub Class_Initialize()
  'Constructor
  
  Set c = New Collection
  dataSizeIndex = 0
  firstDataPosition = 0
  
End Sub
Public Sub enqueue(v As Variant)
  'Add data to the end of the queue
  
  c.Add v, CStr(dataSizeIndex)       'Collection key can only be set to string type
  dataSizeIndex = dataSizeIndex + 1
  
End Sub
Public Function dequeue() As Variant
  'Fetch the data at the beginning of the queue
  
  If c.count = 0 Then
    'Raise error number 1000 when there is no data in the queue
    Err.Raise 1000, "Queue", "There is no data in the queue"
  End If
  
  Dim vType As Long
  vType = VarType(c.Item(CStr(firstDataPosition)))
  
  Select Case vType
    'Use Set syntax for objects, data access objects, and user-defined types
    Case vbObject
      Set dequeue = c.Item(CStr(firstDataPosition))
    Case vbDataObject
      Set dequeue = c.Item(CStr(firstDataPosition))
    Case vbUserDefinedType
      Set dequeue = c.Item(CStr(firstDataPosition))
    Case Else
      dequeue = c.Item(CStr(firstDataPosition))
  End Select
  
  Call c.Remove(CStr(firstDataPosition))
  firstDataPosition = firstDataPosition + 1
  
End Function
Public Function count() As Long
  'Count the number of data in the queue
  
  count = c.count

End Function
Private Sub Class_Terminate()
  'Destructor
  
  Set c = Nothing
  
End Sub

