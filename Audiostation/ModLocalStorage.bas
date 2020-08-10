Attribute VB_Name = "ModLocalStorage"
Public StorageContainer As Nest
Public Sub ClearStorage()
Dim I As Integer

For I = 0 To StorageContainer
    StorageContainer.Remove I
Next
End Sub
Public Sub NewStorageContainer()
Set StorageContainer = New Nest
End Sub
Public Sub AddToStorage(Key As String, Value As String)
StorageContainer.Add Value, Key
End Sub
Public Function GetItemByKey(Key As String, Column As Integer) As String
Dim StoredValue As String
Dim SplitValue

StoredValue = StorageContainer.Item(Key)
SplitValue = Split(StoredValue, ";")

GetItemByKey = SplitValue(Column)
End Function
Public Function GetItemByIndex(Index As Integer, Column As Integer) As String
Dim StoredValue As String
Dim SplitValue

StoredValue = StorageContainer.Item(Index)
SplitValue = Split(StoredValue, ";")

GetItemByIndex = SplitValue(Column)
End Function
