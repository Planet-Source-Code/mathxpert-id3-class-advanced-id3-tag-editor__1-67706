VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiFrameData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BG = "I_"

Dim mCol As Collection

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get Item(ByVal Index As Long) As String
Attribute Item.VB_UserMemId = 0
    Item = mCol(Index)
End Property

Public Property Let Item(ByVal Index As Long, ByVal NewValue As String)
    SetItem mCol, Index, NewValue
End Property

Private Function Cap(ByVal Str As String) As String
    Dim i As Long
    Dim s As String
    s = Trim$(Cap)
    i = InStr(s, Chr$(0))
    If i > 0 Then s = Left$(s, i - 1)
    Cap = s
End Function

Public Sub Add(ByVal Item As String, Optional Before, Optional After, Optional ByVal BinaryMatch As Boolean = False)
    Dim Key As String
    Dim s As String
    If BinaryMatch Then Key = BG & Item Else Key = BG = Cap(Item)
    On Error Resume Next
    s = mCol(Key)
    If Err Then
        Err.Clear
        On Error GoTo 0
        If mCol.Count = 0 Then
            mCol.Add Item, Key
        Else
            If Not IsMissing(Before) Then
                If Before < 1 Then
                    mCol.Add Item, Key, 1
                ElseIf Before <= mCol.Count Then
                    mCol.Add Item, Key, Before
                Else
                    mCol.Add Item, Key
                End If
            ElseIf Not IsMissing(After) Then
                If After < 1 Then
                    mCol.Add Item, Key, , 1
                ElseIf After <= mCol.Count Then
                    mCol.Add Item, Key, , After
                Else
                    mCol.Add Item, Key
                End If
            Else
                mCol.Add Item, Key
            End If
        End If
    End If
End Sub

Public Sub Remove(ByVal Index As Long)
    mCol.Remove Index
End Sub

Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
