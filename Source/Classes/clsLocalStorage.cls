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
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Class:            LocalStorage
' Description:      Used for storing collection data, like playlists, etc.
'
' Date Changed:     05-10-2021
' Date Created:     23-12-2018
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Public IsFilterd As Boolean
Public StorageContainer As New Nest
Private StorageContainerTemp As Nest
Public Function ClearFilter()
Set StorageContainer = StorageContainerTemp
IsFilterd = False
End Function
Public Function Filter(FilterString As String)
Dim i As Integer

Set StorageContainerTemp = New Nest

'First copy the current storage to the temp storage
For i = 1 To StorageContainer.count
    StorageContainerTemp.Add StorageContainer.Item(i), StorageContainer.ItemKey(i)
Next

'Then clear the current storage
StorageContainer.Clear

'Now filter the to temp storage and add the items to the current storage
For i = 1 To StorageContainerTemp.count
    If InStr(1, StorageContainerTemp.Item(i), FilterString, vbTextCompare) > 0 Then
        StorageContainer.Add StorageContainerTemp.Item(i), StorageContainerTemp.ItemKey(i)
    End If
Next

IsFilterd = True
End Function
Public Function IsExistingItem(KeyToFind As String) As Integer
Dim i As Integer

For i = 1 To StorageContainer.count
    If StorageContainer.ItemKey(i) = KeyToFind Then
        IsExistingItem = i
    End If
Next
End Function
Public Sub ListviewToStorage(TargetListview As ListView, KeyColumn As Integer)
Dim i As Integer
Dim C As Integer
Dim StorageItem As String
Dim StorageItemKey As String

For i = 1 To TargetListview.ListItems.count
    StorageItem = vbNullString 'Reset the string builder
    
    For C = 0 To TargetListview.ColumnHeaders.count - 1
        If C = 0 Then
            If C = KeyColumn Then
                StorageItemKey = TargetListview.ListItems(i).Text
            End If
            
            StorageItem = StorageItem & TargetListview.ListItems(i).Text & ";"
        Else
            If C = KeyColumn Then
                StorageItemKey = TargetListview.ListItems(i).SubItems(C)
            End If
            
            StorageItem = StorageItem & TargetListview.ListItems(i).SubItems(C) & ";"
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
Public Sub AddToStorage(key As String, value As String)
StorageContainer.Add value, key
End Sub
Public Function GetItemByKey(key As String, Column As Integer) As String
Dim StoredValue As String
Dim SplitValue

StoredValue = StorageContainer.Item(key)
SplitValue = Split(StoredValue, ";")

GetItemByKey = SplitValue(Column)
End Function
Public Function GetItemByIndex(index As Integer, Column As Integer) As String
Dim StoredValue As String
Dim SplitValue

StoredValue = StorageContainer.Item(index)
SplitValue = Split(StoredValue, ";")

If UBound(SplitValue) = 0 Then
    GetItemByIndex = StoredValue
Else
    GetItemByIndex = SplitValue(Column)
End If
End Function

