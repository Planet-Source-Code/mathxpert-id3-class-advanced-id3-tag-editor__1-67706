Attribute VB_Name = "modUTF8"
Option Explicit

' UTF-8 info @ en.wikipedia.org/wiki/UTF-8

Public Function ANSItoUTF8(ByVal str As String) As String
    Dim i As Long
    Dim CCode As Integer
    Dim s As String
    
    For i = 1 To Len(str)
        CCode = AscW(Mid$(str, i, 1))
        
        Select Case CCode
            Case &H0 To &H7F
                s = s & Chr$(CCode)
            Case &H80 To &H7FF
                s = s & Chr$(CCode \ &H40 Mod &H20 + &HC0) & Chr$(CCode Mod &H40 + &H80)
            Case Else
                s = s & Chr$(CCode \ &H1000 Mod &H10 + &HE0) & Chr$(CCode \ &H40 Mod &H40 + &H80) & Chr$(CCode Mod &H40 + &H80)
        End Select
    Next
    
    ANSItoUTF8 = s
End Function

Public Function UTF8toANSI(ByVal str As String) As String
    Dim i As Long
    Dim CCode As Integer
    Dim CCode2 As Integer
    Dim CCode3 As Integer
    Dim s As String
    
    For i = 1 To Len(str)
        CCode = Asc(Mid$(str, i, 1))
        
        Select Case CCode
            Case &H0 To &HBF
                s = s & ChrW$(CCode)
            Case &HC0 To &HDF
                If i > Len(str) - 1 Then Exit For
                i = i + 1
                CCode2 = Asc(Mid$(str, i, 1))
                s = s & ChrW$((CCode - &HC0) * &H40 + CCode2 - &H80)
            Case Else
                If i > Len(str) - 2 Then Exit For
                i = i + 1
                CCode2 = Asc(Mid$(str, i, 1)): i = i + 1
                CCode3 = Asc(Mid$(str, i, 1))
                s = s & ChrW$((CCode - &HE0) * &H1000 + (CCode2 - &H80) * &H40 + CCode2 - &H80)
        End Select
    Next
    
    UTF8toANSI = s
End Function
