Attribute VB_Name = "modHTML"
Option Explicit

Public Function ToHTML(ByVal Str As String, Optional ByVal UseEntityNames As Boolean = False) As String
    Dim i As Long
    Dim s As String
    Dim a As Integer
    
    s = Str
    If Len(s) = 0 Then Exit Function
    
    i = 1
    
    If UseEntityNames Then
        s = Replace(s, "&", "&amp;")
        s = Replace(s, """", "&quot;")
        s = Replace(s, "'", "&apos;")
        s = Replace(s, "<", "&lt;")
        s = Replace(s, ">", "&gt;")
    Else
        Do
            a = Asc(Mid$(s, i, 1))
            Select Case a
                Case 34, 38, 39, 60, 62, 127, 160
                    s = Left$(s, i - 1) & "&#" & CStr(a) & ";" & Mid$(s, i + 1)
                    i = i + Len("&#" & CStr(a) & ";")
                Case Else
                    i = i + 1
            End Select
        Loop Until i > Len(s)
    End If
    
    ToHTML = s
End Function

Public Function ReplaceHTML(ByVal Str As String) As String
    On Error Resume Next
    
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim t As String
    
    s = Str
    For i = 1 To Len(s)
        If Len(s) - i >= 3 Then
            If Mid$(s, i, 2) = "&#" Then
                j = i + 2
                Do
                    If IsNumeric(Mid$(s, j, 1)) Then
                        j = j + 1
                    Else
                        If j > i + 2 Then
                            j = j - 1
                        End If
                        Exit Do
                    End If
                Loop
                
                If Mid$(s, j + 1, 1) = ";" Then
                    t = Mid$(s, i + 2, j - i - 1)
                    s = Left$(s, i - 1) & Chr$(CLng(t)) & Mid$(s, j + 2)
                End If
            End If
        End If
    Next
    
    s = Replace(s, "&sp;", " ")
    s = Replace(s, "&excl;", "!")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&num;", "#")
    s = Replace(s, "&dollar;", "$")
    s = Replace(s, "&percent;", "%")
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&apos;", "'")
    s = Replace(s, "&lpar;", "(")
    s = Replace(s, "&rpar;", ")")
    s = Replace(s, "&ast;", "*")
    s = Replace(s, "&plus;", "+")
    s = Replace(s, "&comma;", ",")
    s = Replace(s, "&hyphen;", "-")
    s = Replace(s, "&minus;", "-")
    s = Replace(s, "&period;", ".")
    s = Replace(s, "&sol;", "/")
    s = Replace(s, "&colon;", ":")
    s = Replace(s, "&semi;", ";")
    s = Replace(s, "&lt;", "<")
    s = Replace(s, "&equals;", "=")
    s = Replace(s, "&gt;", ">")
    s = Replace(s, "&quest;", "?")
    s = Replace(s, "&commat;", "@")
    s = Replace(s, "&lsqb;", "[")
    s = Replace(s, "&bsol;", "\")
    s = Replace(s, "&rsqb;", "]")
    s = Replace(s, "&circ;", "^")
    s = Replace(s, "&lowbar;", "_")
    s = Replace(s, "&horbar;", "_")
    s = Replace(s, "&grave;", "`")
    s = Replace(s, "&lcub;", "{")
    s = Replace(s, "&verbar;", "|")
    s = Replace(s, "&rcub;", "}")
    s = Replace(s, "&tilde;", "~")
    s = Replace(s, "&lsquor;", Chr$(130))
    s = Replace(s, "&fnof;", Chr$(131))
    s = Replace(s, "&ldquor;", Chr$(132))
    s = Replace(s, "&hellip", Chr$(133))
    s = Replace(s, "&ldots;", Chr$(133))
    s = Replace(s, "&dagger;", Chr$(134))
    s = Replace(s, "&Dagger;", Chr$(135))
    s = Replace(s, "&permil;", Chr$(137))
    s = Replace(s, "&Scaron;", Chr$(138))
    s = Replace(s, "&lsaquo;", Chr$(139))
    s = Replace(s, "&OElig;", Chr$(140))
    s = Replace(s, "&lsquo;", Chr$(145))
    s = Replace(s, "&rsquor;", Chr$(145))
    s = Replace(s, "&rsquo;", Chr$(146))
    s = Replace(s, "&ldquo;", Chr$(147))
    s = Replace(s, "&rdquor;", Chr$(148))
    s = Replace(s, "&bull;", Chr$(149))
    s = Replace(s, "&ndash;", Chr$(150))
    s = Replace(s, "&endash;", Chr$(150))
    s = Replace(s, "&mdash;", Chr$(151))
    s = Replace(s, "&emdash;", Chr$(151))
    s = Replace(s, "&trade;", Chr$(153))
    s = Replace(s, "&scaron;", Chr$(154))
    s = Replace(s, "&rsaquo;", Chr$(155))
    s = Replace(s, "&oelig;", Chr$(156))
    s = Replace(s, "&Yuml;", Chr$(159))
    s = Replace(s, "&nbsp;", " ")
    s = Replace(s, "&iexcl;", Chr$(161))
    s = Replace(s, "&cent;", Chr$(162))
    s = Replace(s, "&pound;", Chr$(163))
    s = Replace(s, "&curren;", Chr$(164))
    s = Replace(s, "&yen;", Chr$(165))
    s = Replace(s, "&brvbar;", Chr$(166))
    s = Replace(s, "&brkbar;", Chr$(166))
    s = Replace(s, "&sect;", Chr$(167))
    s = Replace(s, "&uml;", Chr$(168))
    s = Replace(s, "&die;", Chr$(168))
    s = Replace(s, "&copy;", Chr$(169))
    s = Replace(s, "&ordf;", Chr$(170))
    s = Replace(s, "&laquo;", Chr$(171))
    s = Replace(s, "&not;", Chr$(172))
    s = Replace(s, "&shy;", Chr$(173))
    s = Replace(s, "&reg;", Chr$(174))
    s = Replace(s, "&macr;", Chr$(175))
    s = Replace(s, "&hibar;", Chr$(175))
    s = Replace(s, "&deg;", Chr$(176))
    s = Replace(s, "&plusmn;", Chr$(177))
    s = Replace(s, "&sup2;", Chr$(178))
    s = Replace(s, "&sup3;", Chr$(179))
    s = Replace(s, "&acute;", Chr$(180))
    s = Replace(s, "&micro;", Chr$(181))
    s = Replace(s, "&para;", Chr$(182))
    s = Replace(s, "&middot;", Chr$(183))
    s = Replace(s, "&cedil;", Chr$(184))
    s = Replace(s, "&sup1;", Chr$(185))
    s = Replace(s, "&ordm;", Chr$(186))
    s = Replace(s, "&raquo;", Chr$(187))
    s = Replace(s, "&frac14;", Chr$(188))
    s = Replace(s, "&frac12;", Chr$(189))
    s = Replace(s, "&half;", Chr$(189))
    s = Replace(s, "&frac34;", Chr$(190))
    s = Replace(s, "&iquest;", Chr$(191))
    s = Replace(s, "&Agrave;", Chr$(192))
    s = Replace(s, "&Aacute;", Chr$(193))
    s = Replace(s, "&Acirc;", Chr$(194))
    s = Replace(s, "&Atilde;", Chr$(195))
    s = Replace(s, "&Auml;", Chr$(196))
    s = Replace(s, "&Aring;", Chr$(197))
    s = Replace(s, "&AElig;", Chr$(198))
    s = Replace(s, "&Ccedil;", Chr$(199))
    s = Replace(s, "&Egrave;", Chr$(200))
    s = Replace(s, "&Eacute;", Chr$(201))
    s = Replace(s, "&Ecirc;", Chr$(202))
    s = Replace(s, "&Euml;", Chr$(203))
    s = Replace(s, "&Igrave;", Chr$(204))
    s = Replace(s, "&Iacute;", Chr$(205))
    s = Replace(s, "&Icirc;", Chr$(206))
    s = Replace(s, "&Iuml;", Chr$(207))
    s = Replace(s, "&ETH;", Chr$(208))
    s = Replace(s, "&Ntilde;", Chr$(209))
    s = Replace(s, "&Ograve;", Chr$(210))
    s = Replace(s, "&Oacute;", Chr$(211))
    s = Replace(s, "&Ocirc;", Chr$(212))
    s = Replace(s, "&Otilde;", Chr$(213))
    s = Replace(s, "&Ouml;", Chr$(214))
    s = Replace(s, "&times;", Chr$(215))
    s = Replace(s, "&Oslash;", Chr$(216))
    s = Replace(s, "&Ugrave;", Chr$(217))
    s = Replace(s, "&Uacute;", Chr$(218))
    s = Replace(s, "&Ucirc;", Chr$(219))
    s = Replace(s, "&Uuml;", Chr$(220))
    s = Replace(s, "&Yacute;", Chr$(221))
    s = Replace(s, "&THORN;", Chr$(222))
    s = Replace(s, "&szlig;", Chr$(223))
    s = Replace(s, "&agrave;", Chr$(224))
    s = Replace(s, "&aacute;", Chr$(225))
    s = Replace(s, "&acirc;", Chr$(226))
    s = Replace(s, "&atilde;", Chr$(227))
    s = Replace(s, "&auml;", Chr$(228))
    s = Replace(s, "&aring;", Chr$(229))
    s = Replace(s, "&aelig;", Chr$(230))
    s = Replace(s, "&ccedil;", Chr$(231))
    s = Replace(s, "&egrave;", Chr$(232))
    s = Replace(s, "&eacute;", Chr$(233))
    s = Replace(s, "&ecirc;", Chr$(234))
    s = Replace(s, "&euml;", Chr$(235))
    s = Replace(s, "&igrave;", Chr$(236))
    s = Replace(s, "&iacute;", Chr$(237))
    s = Replace(s, "&icirc;", Chr$(238))
    s = Replace(s, "&iuml;", Chr$(239))
    s = Replace(s, "&eth;", Chr$(240))
    s = Replace(s, "&ntilde;", Chr$(241))
    s = Replace(s, "&ograve;", Chr$(242))
    s = Replace(s, "&oacute;", Chr$(243))
    s = Replace(s, "&ocirc;", Chr$(244))
    s = Replace(s, "&otilde;", Chr$(245))
    s = Replace(s, "&ouml;", Chr$(246))
    s = Replace(s, "&divide;", Chr$(247))
    s = Replace(s, "&oslash;", Chr$(248))
    s = Replace(s, "&ugrave;", Chr$(249))
    s = Replace(s, "&uacute;", Chr$(250))
    s = Replace(s, "&ucirc;", Chr$(251))
    s = Replace(s, "&uuml;", Chr$(252))
    s = Replace(s, "&yacute;", Chr$(253))
    s = Replace(s, "&thorn;", Chr$(254))
    s = Replace(s, "&yuml;", Chr$(255))
    
    ReplaceHTML = s
End Function

