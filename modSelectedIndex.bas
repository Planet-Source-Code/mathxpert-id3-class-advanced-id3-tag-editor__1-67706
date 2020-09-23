Attribute VB_Name = "modSelectedIndex"
Option Explicit

Public Function SelectedIndex(ListView As ListView) As Long
    On Error GoTo NotSelected
    SelectedIndex = ListView.SelectedItem.Index
    Exit Function
    
NotSelected:
    SelectedIndex = -1
End Function
