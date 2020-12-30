# VBA Queue

[![license](https://img.shields.io/github/license/mashape/apistatus.svg)](LICENSE)

This cls file implementing queue with vba collection object.

## Install
1. Download Queue.cls.
2. Add Queue.cls to class module of the VBA Project.

## Usage
1. Make Queue Object From Queue.cls.
2. Use Enqueue to add data to the queue.
3. Use Dequeue to pop data from the queue.
4. If you use Dequeue when there is no data in the queue, Raise error 1000.


## Code Example
```VBA
Public Sub TestQueue()

On Error GoTo ErrorLabel
  
  Dim q As Queue
  Set q = New Queue
  
  q.Enqueue 1
  q.Enqueue 2
  q.Enqueue 3
  MsgBox q.Dequeue
  MsgBox q.Dequeue
  MsgBox q.Dequeue
  MsgBox q.Dequeue  'Raise error 1000

ErrorLabel:
  If Err.Number = 1000 Then
    MsgBox Err.Description, vbCritical
  End If
  
End Sub
```

>Data retrieved using Dequeue also remains in memory, so it is not suitable for handling large amounts of data.
