VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LocalStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsFilterd As Boolean
Public StorageContainer As New Nest
Private StorageContainerTemp As Nest
Public Function ClearFilter()
Set StorageContainer = StorageContainerTemp
IsFilterd = False
End Function
Public Function Filter(FilterString As String)
Dim I As Integer

Set StorageContainerTemp = New Nest

'First copy the current storage to the temp storage
For I = 1 To StorageContainer.Count
    StorageContainerTemp.Add StorageContainer.Item(I), StorageContainer.ItemKey(I)
Next

'Then clear the current storage
StorageContainer.Clear

'Now filter the to temp storage and add the items to the current storage
For I = 1 To StorageContainerTemp.Count
    If InStr(1, StorageContainerTemp.Item(I), FilterString, vbTextCompare) > 0 Then
        StorageContainer.Add StorageContainerTemp.Item(I), StorageContainerTemp.ItemKey(I)
    End If
Next

IsFilterd = True
End Function
Public Function IsExistingItem(KeyToFind As String) As Integer
Dim I As Integer

For I = 1 To StorageContainer.Count
    If StorageContainer.ItemKey(I) = KeyToFind Then
        IsExistingItem = I
    End If
Next
End Function
Public Sub ListviewToStorage(TargetListview As ListView, KeyColumn As Integer)
Dim I As Integer
Dim C As Integer
Dim StorageItem As String
Dim StorageItemKey As String

For I = 1 To TargetListview.ListItems.Count
    StorageItem = vbNullString 'Reset the string builder
    
    For C = 0 To TargetListview.ColumnHeaders.Count - 1
        If C = 0 Then
            If C = KeyColumn Then
                StorageItemKey = TargetListview.ListItems(I).Text
            End If
            
            StorageItem = StorageItem & TargetListview.ListItems(I).Text & ";"
        Else
            If C = KeyColumn Then
                StorageItemKey = TargetListview.ListItems(I).SubItems(C)
            End If
            
            StorageItem = StorageItem & TargetListview.ListItems(I).SubItems(C) & ";"
        End If
    Next
    
    'Remove the last ;
    StorageItem = Left(StorageItem, Len(StorageItem) - 1)
     
    Me.AddToStorage StorageItemKey, StorageItem
Next
End Sub
Public Sub ClearStorage()
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

If UBound(SplitValue) = 0 Then
    GetItemByIndex = StoredValue
Else
    GetItemByIndex = SplitValue(Column)
End If
End Function

