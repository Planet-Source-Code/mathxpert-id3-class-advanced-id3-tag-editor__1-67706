Attribute VB_Name = "mLVSort"
'Make sure the module name is
'mLVSort.  Set this in
'your properties window

Option Explicit
Public objFind As LV_FINDINFO
Public objItem As LV_ITEM
Public Resort As Boolean

'variable to hold the sort order (ascending or descending)
Public sOrder As Boolean
'variable to hold sort column
Public sColumn As Long
'variable to hold the sort level (used in multi-column sorting)
Public sLevel As Long

Public Type POINT
  x As Long
  Y As Long
End Type

Public Type LV_FINDINFO
  flags As Long
  psz As String
  lParam As Long
  pt As POINT
  vkDirection As Long
End Type

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
 
'Constants
Public Const LVFI_PARAM = 1
Public Const LVIF_TEXT = &H1

Public Const LVM_FIRST = &H1000
Public Const LVM_FINDITEM = LVM_FIRST + 13
Public Const LVM_GETITEMTEXT = LVM_FIRST + 45
Public Const LVM_SORTITEMS = LVM_FIRST + 48
     
'API declarations
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
  ByVal hWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long) As Long

Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" ( _
  ByVal hWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long

Private Const LVM_GETNEXTITEM As Long = (LVM_FIRST + 12)
Private Const LVNI_FOCUSED As Long = &H1
Public Const LVM_ENSUREVISIBLE As Long = (LVM_FIRST + 19)
Public Const LVM_UPDATE As Long = (LVM_FIRST + 42)

' This function plays a major role in multi-column sorting (iTunes does this in a similar fashion)
' The level starts out at 0 (the primary, most important level - highest priority) and goes all the way to 10 (the least important level - lowest priority)
' All the numbers inside the array functions indicate column indices (indices are zero-based by which 0 indicates the very first column), and you can decide how to sort the priorities if you want to (simply rearrange the column indices in the array functions)
Private Function SubsortKey(ByVal Level As Long) As Long
    Dim sChain() As Variant
    
    Select Case sColumn ' Determine which column index is to be sorted first
        Case 0 ' Filename column
            sChain = Array(0, 2, 3, 5, 6, 1, 4, 7, 10, 8, 9) ' A RHETORICAL QUESTION: Would multi-column sorting ever be needed if the Filename column is set as the primary sort column? (... as there can only be one instance of a filename per directory.)
        Case 1 ' Title column
            sChain = Array(1, 2, 3, 5, 6, 4, 7, 10, 0, 8, 9)
        Case 2 ' Artist column
            sChain = Array(2, 3, 5, 6, 1, 4, 7, 10, 0, 8, 9)
        Case 3 ' Album column
            sChain = Array(3, 5, 6, 2, 1, 4, 7, 10, 0, 8, 9)
        Case 4 ' Genre column
            sChain = Array(4, 2, 3, 5, 6, 1, 7, 10, 0, 8, 9)
        Case 5 ' Track number column
            sChain = Array(5, 6, 2, 3, 1, 4, 7, 10, 0, 8, 9)
        Case 6 ' Tracks total column
            sChain = Array(6, 5, 2, 3, 1, 4, 7, 10, 0, 8, 9)
        Case 7 ' Year column
            sChain = Array(7, 2, 3, 5, 6, 1, 4, 10, 0, 8, 9)
        Case 8 ' Duration column
            sChain = Array(8, 2, 3, 5, 6, 1, 4, 7, 10, 0, 9)
        Case 9 ' Bit rate column
            sChain = Array(9, 2, 3, 5, 6, 1, 4, 7, 10, 0, 8)
        Case 10 ' Comments column
            sChain = Array(10, 3, 5, 6, 2, 1, 4, 7, 0, 8, 9)
    End Select
    
    SubsortKey = sChain(Level)
End Function

' This procedure allows alphanumeric strings to be sorted properly
Private Sub PatternMatching(ByVal Input1 As String, ByVal Input2 As String, Output1 As String, Output2 As String)
    Dim i As Long
    Dim j0 As Long
    Dim j1 As Long
    Dim finalJ0 As Long
    Dim finalJ1 As Long
    Dim k0 As Long
    Dim k1 As Long
    Dim lastK0 As Long
    Dim lastK1 As Long
    Dim x As String
    Dim Num1 As String
    Dim Num2 As String
    
    Output1 = Input1
    Output2 = Input2
      
    k0 = 1
    k1 = 1
    lastK0 = 1
    lastK1 = 1
    
    Do
        finalJ0 = 0
        finalJ1 = 0
        
        For i = 0 To 9
            j0 = InStr(k0, Output1, CStr(i))
            j1 = InStr(k1, Output2, CStr(i))
            
            If j0 > 0 Then
                If finalJ0 = 0 Then
                    finalJ0 = j0
                Else
                    If finalJ0 > j0 Then finalJ0 = j0
                End If
            End If
            
            If j1 > 0 Then
                If finalJ1 = 0 Then
                    finalJ1 = j1
                Else
                    If finalJ1 > j1 Then finalJ1 = j1
                End If
            End If
        Next
        
        ' If the numbers don't start in the same place, exit
        If finalJ0 = 0 Or finalJ1 = 0 Or finalJ0 <> finalJ1 Then
            Exit Do
        End If
        
        ' If the strings in between the numbers or up to the numbers don't match, exit
        If Mid$(Output1, lastK0, finalJ0 - lastK0) <> Mid$(Output2, lastK1, finalJ1 - lastK1) Then
            Exit Do
        End If
        
        Num1 = ""
        k0 = finalJ0
        Do
            x = Mid$(Output1, k0, 1)
            If IsNumeric(x) Then
                Num1 = Num1 & x
                k0 = k0 + 1
            Else
                Exit Do
            End If
        Loop
        
        Num2 = ""
        k1 = finalJ1
        Do
            x = Mid$(Output2, k1, 1)
            If IsNumeric(x) Then
                Num2 = Num2 & x
                k1 = k1 + 1
            Else
                Exit Do
            End If
        Loop
        
        ' Add leading zeros to the lesser number to sort the alphanumeric strings properly
        If k0 < k1 Then
            Num1 = Format$(Num1, String$(Len(Num2), "0"))
            Output1 = Left$(Output1, finalJ0 - 1) & Num1 & Mid$(Output1, k0)
            k0 = k1
        ElseIf k0 > k1 Then
            Num2 = Format$(Num2, String$(Len(Num1), "0"))
            Output2 = Left$(Output2, finalJ1 - 1) & Num2 & Mid$(Output2, k1)
            k1 = k0
        End If
        
        lastK0 = k0
        lastK1 = k1
    Loop
End Sub

Private Function DurationToLong(ByVal sDuration As String, bSuccess As Boolean) As Long
    Dim aDuration() As String
    Dim i As Long
    Dim lB As Long
    Dim uB As Long
    Dim num As Long
    Dim Multiplier As Long
    
    aDuration = Split(sDuration, ":")
    Multiplier = 1
    bSuccess = False
    
    On Error Resume Next
    
    lB = LBound(aDuration)
    uB = UBound(aDuration)
    
    If Err Then Exit Function
    On Error GoTo 0
    
    If lB < uB - 2 Then
        lB = uB - 2
    End If
    
    For i = uB To lB Step -1
        If IsNumeric(aDuration(i)) Then
            If i = uB - 2 Then
                If aDuration(i) >= 0 Then
                    bSuccess = True
                End If
            ElseIf i > uB - 2 Then
                If aDuration(i) >= 0 And aDuration(i) < 60 Then
                    bSuccess = True
                End If
            End If
        End If
        If Not bSuccess Then Exit Function
        
        num = num + CLng(aDuration(i)) * Multiplier
        Multiplier = Multiplier * 60
    Next
    
    DurationToLong = num
End Function

Private Function BitRateToLong(ByVal sBitRate As String, bSuccess As Boolean) As Long
    Dim i As Long
    Dim x As String
    Dim dX As Long
    Dim Y As String
    
    bSuccess = False
    i = InStr(sBitRate, " kbps ")
    If i > 0 Then
        x = Left$(sBitRate, i - 1)
        If IsNumeric(x) Then
            If CDbl(x) > 0 And CDbl(x) < 214748364 Then
                dX = CLng(x) * 10
                Y = Mid$(sBitRate, i + 6)
                If Y = "VBR" Then dX = dX + 1
                If Y = "CBR" Or Y = "VBR" Then
                    bSuccess = True
                    BitRateToLong = dX
                End If
            End If
        End If
    End If
End Function

Public Function CompareValues(ByVal lParam1 As Long, _
                              ByVal lParam2 As Long, _
                              ByVal hWnd As Long) As Long
     
'CompareValues: This is the sorting routine that gets passed to the
'ListView control to provide the comparison test for numeric values.

  'Compare returns:
  ' -1 = Less Than
  ' 0 = Equal
  ' 1 = Greater Than
  
Dim val1 As String, val2 As String, dVal1 As Double, dVal2 As Double, dtVal1 As Date, dtVal2 As Date, sVal1 As String, sVal2 As String, bNum As Boolean, bDate As Boolean
On Error GoTo CDERR
    'Obtain the item names and values corresponding
    'to the input parameters
    bNum = False
    bDate = False
    val1 = ListView_GetItemValueStr(hWnd, lParam1, SubsortKey(sLevel))
    val2 = ListView_GetItemValueStr(hWnd, lParam2, SubsortKey(sLevel))
    
    If IsNumeric(val1) And IsNumeric(val2) Then
        bNum = True
        dVal1 = CDbl(val1)
        dVal2 = CDbl(val2)
    ElseIf IsDate(val1) And IsDate(val2) Then
        bDate = True
        dtVal1 = CDate(val1)
        dtVal2 = CDate(val2)
    Else
        sVal1 = LCase$(val1)
        sVal2 = LCase$(val2)
        PatternMatching sVal1, sVal2, sVal1, sVal2
    End If
    
    'based on the Public variable sOrder set in the
    'columnheader click sub, sort the values appropriately:
    Select Case sOrder
        Case True:    'sort descending
            
            If bNum Then
                If dVal1 < dVal2 Then
                    sLevel = 0
                    CompareValues = -1
                ElseIf dVal1 = dVal2 Then
                    If sLevel = 10 Then
                        sLevel = 0
                        CompareValues = 0
                    Else
                        sLevel = sLevel + 1
                        CompareValues = CompareValues(lParam1, lParam2, hWnd)
                    End If
                Else
                    sLevel = 0
                    CompareValues = 1
                End If
            ElseIf bDate Then
                If dtVal1 < dtVal2 Then
                    sLevel = 0
                    CompareValues = -1
                ElseIf dtVal1 = dtVal2 Then
                    If sLevel = 10 Then
                        sLevel = 0
                        CompareValues = 0
                    Else
                        sLevel = sLevel + 1
                        CompareValues = CompareValues(lParam1, lParam2, hWnd)
                    End If
                Else
                    sLevel = 0
                    CompareValues = 1
                End If
            Else
                If sVal1 < sVal2 Then
                    sLevel = 0
                    CompareValues = -1
                ElseIf sVal1 = sVal2 Then
                    If sLevel = 10 Then
                        sLevel = 0
                        CompareValues = 0
                    Else
                        sLevel = sLevel + 1
                        CompareValues = CompareValues(lParam1, lParam2, hWnd)
                    End If
                Else
                    sLevel = 0
                    CompareValues = 1
                End If
            End If
      
        Case Else: 'sort ascending
   
            If bNum Then
                If dVal1 > dVal2 Then
                    sLevel = 0
                    CompareValues = -1
                ElseIf dVal1 = dVal2 Then
                    If sLevel = 10 Then
                        sLevel = 0
                        CompareValues = 0
                    Else
                        sLevel = sLevel + 1
                        CompareValues = CompareValues(lParam1, lParam2, hWnd)
                    End If
                Else
                    sLevel = 0
                    CompareValues = 1
                End If
            ElseIf bDate Then
                If dtVal1 > dtVal2 Then
                    sLevel = 0
                    CompareValues = -1
                ElseIf dtVal1 = dtVal2 Then
                    If sLevel = 10 Then
                        sLevel = 0
                        CompareValues = 0
                    Else
                        sLevel = sLevel + 1
                        CompareValues = CompareValues(lParam1, lParam2, hWnd)
                    End If
                Else
                    sLevel = 0
                    CompareValues = 1
                End If
            Else
                If sVal1 > sVal2 Then
                    sLevel = 0
                    CompareValues = -1
                ElseIf sVal1 = sVal2 Then
                    If sLevel = 10 Then
                        sLevel = 0
                        CompareValues = 0
                    Else
                        sLevel = sLevel + 1
                        CompareValues = CompareValues(lParam1, lParam2, hWnd)
                    End If
                Else
                    sLevel = 0
                    CompareValues = 1
                End If
            End If
   
    End Select
    Exit Function
CDERR:
    If sLevel = 10 Then
        sLevel = 0
        CompareValues = 0
    Else
        sLevel = sLevel + 1
        CompareValues = CompareValues(lParam1, lParam2, hWnd)
    End If
End Function

Public Function ListView_GetItemValueStr(hWnd As Long, lParam As Long, Optional ColumnIndex) As String
'The optional ColumnIndex argument WILL be needed if we want to perform multi-column sorting
Dim r As Long, hIndex As Long, s As String, b As Boolean, l As Long
    'Convert the input parameter to an index in the list view
    objFind.flags = LVFI_PARAM
    objFind.lParam = lParam
    hIndex = SendMessageAny(hWnd, LVM_FINDITEM, -1, objFind)
     
    'Obtain the value of the specified list view item.
    'The objItem.iSubItem member is set to the index
    'of the column that is being retrieved.
    objItem.mask = LVIF_TEXT
    If IsMissing(ColumnIndex) Then
        objItem.iSubItem = sColumn
    Else
        objItem.iSubItem = ColumnIndex
    End If
    objItem.pszText = Space$(256)
    objItem.cchTextMax = Len(objItem.pszText)
     
    'get the string at subitem 2
    r = SendMessageAny(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
     
    'and convert it into a long
    If r > 0 Then
        s = Left$(objItem.pszText, r)
        l = DurationToLong(s, b)
        If b Then
            ListView_GetItemValueStr = CStr(l)
        Else
            l = BitRateToLong(s, b)
            If b Then
                ListView_GetItemValueStr = CStr(l)
            Else
                ListView_GetItemValueStr = s
            End If
        End If
    End If
End Function

Public Sub SortLvwOnLong(lvw As ListView, ColIndex As Long)
    lvw.Sorted = False
    If lvw.SortKey = ColIndex - 1 And Not Resort Then
        If lvw.SortOrder = lvwAscending Then
            lvw.SortOrder = lvwDescending
        Else
            lvw.SortOrder = lvwAscending
        End If
    ElseIf Not Resort Then
        lvw.SortKey = ColIndex - 1
        lvw.SortOrder = lvwAscending
    End If
    sColumn = ColIndex - 1
    sOrder = (lvw.SortOrder = lvwAscending)
    sLevel = 0
    SendMessageLong lvw.hWnd, LVM_SORTITEMS, lvw.hWnd, AddressOf CompareValues
End Sub

Public Function SelectedItemIdx(lvw As ListView) As Long
    SelectedItemIdx = SendMessageLong(lvw.hWnd, LVM_GETNEXTITEM, -1, LVNI_FOCUSED)
End Function

Public Sub APIEnsureVisible(lvw As ListView, ByVal Index As Long, Optional ByVal EnsureUpdated As Boolean = False)
    If Index <> -1 Then
        SendMessageLong lvw.hWnd, LVM_ENSUREVISIBLE, Index, 0&
        If EnsureUpdated Then
            SendMessageLong lvw.hWnd, LVM_UPDATE, Index, 0&
        End If
    End If
End Sub

Public Sub EnsureSelVisible(lvw As ListView, Optional ByVal EnsureUpdated As Boolean = False)
    APIEnsureVisible lvw, SelectedItemIdx(lvw), EnsureUpdated
End Sub
