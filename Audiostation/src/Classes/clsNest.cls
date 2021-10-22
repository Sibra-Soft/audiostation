VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Nest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' /////////////////////////////////////////////////////////////////////////////////
' Class:            Nest
' Description:      Extended version of the dictonary build-in dictonary class
'
' Date Changed:     05-10-2021
' Date Created:     15-11-2020
' Author:           Sibra-Soft - Alex van den Berg
' /////////////////////////////////////////////////////////////////////////////////

Private mCol As Collection
Public Function Clear()
Set mCol = New Collection
End Function
Public Function Exists(Optional sItem As String, Optional sKey As String) As Boolean
Dim i As Integer
Dim ReturnResult As Boolean

ReturnResult = False

If sItem = vbNullString And sKey = vbNullString Then
    Call Err.Raise(9, "Class - Nest", "You must specify a item or key name to find")
    Exit Function
End If

If sItem <> vbNullString And sKey <> vbNullString Then
    Call Err.Raise(10, "Class - Nest", "You can only find one of the two options")
    Exit Function
End If

For i = 1 To mCol.count
    'Find item
    If Me.Item(i) = sItem And sItem <> vbNullString Then
        ReturnResult = True
        
        Exists = ReturnResult
        Exit Function
    End If
    
    'Find key
    If Me.ItemKey(i) = sKey And sKey <> vbNullString Then
        ReturnResult = True
        
        Exists = ReturnResult
        Exit Function
    End If
Next

Exists = ReturnResult
End Function
Public Sub Add(Item As Variant, Optional sKey As String, Optional Before As Variant, Optional After As Variant)
On Error GoTo ErrorHandler
Dim objNewMember() As Variant
ReDim objNewMember(0 To 1)

If IsObject(Item) Then
    Set objNewMember(0) = Item
Else
    objNewMember(0) = Item
End If

objNewMember(1) = sKey

If Len(sKey) = 0 Then
    mCol.Add objNewMember, , Before, After
Else
    mCol.Add objNewMember, sKey, Before, After
End If

ErrorHandler:
Select Case Err.Number
    Case 0
    Case Else
End Select
End Sub

Public Property Get Item(vntIndexKey As Variant) As Variant
   Dim vItem() As Variant
   vItem = mCol(vntIndexKey)
   
   If IsObject(vItem(0)) Then
       Set Item = vItem(0)
   Else
       Item = vItem(0)
   End If
End Property
Public Property Get ItemKey(vntIndexKey As Variant) As String
   ItemKey = mCol(vntIndexKey)(1)
End Property

Public Property Get count() As Long
   'used when retrieving the number of elements in the
   'collection. Syntax: Debug.Print x.Count
   count = mCol.count
End Property

Public Sub Remove(vntIndexKey As Variant)
   'used when removing an element from the collection
   'vntIndexKey contains either the Index or Key, which is why
   'it is declared as a Variant
   'Syntax: x.Remove(xyz)
   
   mCol.Remove vntIndexKey
End Sub
Private Sub Class_Initialize()
   'creates the collection when this class is created
   Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
   'destroys collection when this class is terminated
   Set mCol = Nothing
End Sub
