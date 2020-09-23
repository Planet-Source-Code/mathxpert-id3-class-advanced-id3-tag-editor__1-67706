Attribute VB_Name = "modCollection"
Option Explicit

Public Sub SetItem(ByVal Col As Collection, ByVal Index As Long, Item)
    Dim bSetBefore As Boolean
    
    bSetBefore = (Col.Count > Index And Col.Count > 1)
    Col.Remove Index
    If bSetBefore Then
        Col.Add Item, Before:=Index
    Else
        Col.Add Item
    End If
End Sub
