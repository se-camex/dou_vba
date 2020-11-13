Sub calibri9()
    ' converte fonte para calibri 9
    Selection.WholeStory
    Selection.Font.Name = "Calibri"
    Selection.Font.Size = 9
End Sub

Sub converter_tabela_12()
    '
    ' converter_tabela_12 Macro
    ' converte tabela para rtf 12cm ou 25cm
    Dim iTblWidth As Integer
    Dim iCount As Integer
    For Each oTable In ActiveDocument.Tables
        With oTable
            oTable.Select
            Selection.Font.Name = "Calibri"
            Selection.Font.Grow
            Selection.Font.Size = 9
            Selection.Tables(1).Style = "Tabela com grade"
            With Selection.ParagraphFormat
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceSingle
                .WidowControl = False
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .NoLineNumber = False
                .Hyphenation = True
                .FirstLineIndent = CentimetersToPoints(0)
                .OutlineLevel = wdOutlineLevelBodyText
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .MirrorIndents = False
                .TextboxTightWrap = wdTightNone
                .CollapsedByDefault = False
            End With
        End With
        oTable.PreferredWidthType = wdPreferredWidthPoints
        oTable.PreferredWidth = CentimetersToPoints(12)
        'For iCount = 1 To oTable.Columns.Count
        '            iTblWidth = iTblWidth + oTable.Columns(iCount).Width
        '        Next iCount
        '    If iTblWidth > CentimetersToPoints(12.1) Then
        '        'MsgBox CentimetersToPoints(12)
        '        oTable.PreferredWidth = CentimetersToPoints(25)
        '    End If
        '    iTblWidth = 0
        If oTable.PreferredWidth > CentimetersToPoints(12.1) And oTable.PreferredWidth < CentimetersToPoints(50) Then
            oTable.PreferredWidth = CentimetersToPoints(25)
        End If
    Next oTable
End Sub


Sub converter_tabela_25()
    '
    ' converter_tabela_25 Macro
    ' converte tabela para rtf 12cm ou 25cm
    Dim iTblWidth As Integer
    Dim iCount As Integer
    For Each oTable In ActiveDocument.Tables
        With oTable
            oTable.Select
            Selection.Font.Name = "Calibri"
            Selection.Font.Grow
            Selection.Font.Size = 9
            Selection.Tables(1).Style = "Tabela com grade"
            With Selection.ParagraphFormat
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceSingle
                .WidowControl = False
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .NoLineNumber = False
                .Hyphenation = True
                .FirstLineIndent = CentimetersToPoints(0)
                .OutlineLevel = wdOutlineLevelBodyText
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .MirrorIndents = False
                .TextboxTightWrap = wdTightNone
                .CollapsedByDefault = False
            End With
        End With
        oTable.PreferredWidthType = wdPreferredWidthPoints
        oTable.PreferredWidth = CentimetersToPoints(25)
    Next oTable
End Sub

Sub converte_rodape()
    Dim i As Long, RngNt As Range, RngTxt As Range
    With ActiveDocument
        For i = .Footnotes.Count To 1 Step -1
            With .Footnotes(i)
                Set RngNt = .Range
                With RngNt
                    .End = .End
                    .Start = .Start
                End With
                Set RngTxt = .Reference
                With RngTxt
                    .InsertAfter " [nota " & i & "] "
                    .Collapse wdCollapseEnd
                    .InsertAfter " [\nota " & i & "] "
                    .Collapse wdCollapseStart
                    .FormattedText = RngNt.FormattedText
                End With
                .Delete
            End With
        Next
    End With
End Sub

Sub nova_linha_tabela()
    '
    ' nova_linha_tabela Macro
    ' Insere nova linha na tabela e formata como caixa
    '
    Selection.InsertRowsBelow 1
    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
End Sub

Sub SymbolToUnicode()
    ' ver também https://www.stat.auckland.ac.nz/~paul/R/CM/AdobeSym.html
    ' Source https://groups.google.com/forum/#!topic/microsoft.public.word.printingfonts/pPtEuYqrlvQ
    ' Select document or range before running the macro
    Dim myFont As String
    Dim myCharNum As Long
    Dim myRange As Range
    Dim myChar As Range
    Dim i As Long, CharCount As Long
    Set myRange = Selection.Range.Duplicate
    CharCount = myRange.ComputeStatistics(wdStatisticCharacters)
    For Each myChar In myRange.Characters
        i = i + 1
        If myChar.Font.Name = "Symbol" Then
            myCharNum = AscW(myChar.Text) And &HFFFF&
            ' Decorative Fonts are mapped to a
            ' "private use" code page starting at &HF000
            myCharNum = myCharNum - &HF000&
            myChar.Font.Name = myChar.Style.Font.Name
            Select Case myCharNum
            Case &H22                            ' # FOR ALL
                myChar.Text = ChrW(&H2200)
            Case &H24                            ' # THERE EXISTS
                myChar.Text = ChrW(&H2203)
            Case &H27                            ' # CONTAINS AS MEMBER
                myChar.Text = ChrW(&H220B)
            Case &H2A                            ' # ASTERISK OPERATOR
                myChar.Text = ChrW(&H2217)
            Case &H2D                            ' # MINUS SIGN
                myChar.Text = ChrW(&H2212)
            Case &H40                            ' # APPROXIMATELY EQUAL TO
                myChar.Text = ChrW(&H2245)
            Case &H41                            ' # GREEK CAPITAL LETTER ALPHA
                myChar.Text = ChrW(&H391)
            Case &H42                            ' # GREEK CAPITAL LETTER BETA
                myChar.Text = ChrW(&H392)
            Case &H43                            ' # GREEK CAPITAL LETTER CHI
                myChar.Text = ChrW(&H3A7)
            Case &H44                            ' # GREEK CAPITAL LETTER DELTA
                myChar.Text = ChrW(&H394)
            Case &H44                            ' # INCREMENT
                myChar.Text = ChrW(&H2206)
            Case &H45                            ' # GREEK CAPITAL LETTER EPSILON
                myChar.Text = ChrW(&H395)
            Case &H46                            ' # GREEK CAPITAL LETTER PHI
                myChar.Text = ChrW(&H3A6)
            Case &H47                            ' # GREEK CAPITAL LETTER GAMMA
                myChar.Text = ChrW(&H393)
            Case &H48                            ' # GREEK CAPITAL LETTER ETA
                myChar.Text = ChrW(&H397)
            Case &H49                            ' # GREEK CAPITAL LETTER IOTA
                myChar.Text = ChrW(&H399)
            Case &H4A                            ' # GREEK THETA SYMBOL
                myChar.Text = ChrW(&H3D1)
            Case &H4B                            ' # GREEK CAPITAL LETTER KAPPA
                myChar.Text = ChrW(&H39A)
            Case &H4C                            ' # GREEK CAPITAL LETTER LAMDA
                myChar.Text = ChrW(&H39B)
            Case &H4D                            ' # GREEK CAPITAL LETTER MU
                myChar.Text = ChrW(&H39C)
            Case &H4E                            ' # GREEK CAPITAL LETTER NU
                myChar.Text = ChrW(&H39D)
            Case &H4F                            ' # GREEK CAPITAL LETTER OMICRON
                myChar.Text = ChrW(&H39F)
            Case &H50                            ' # GREEK CAPITAL LETTER PI
                myChar.Text = ChrW(&H3A0)
            Case &H51                            ' # GREEK CAPITAL LETTER THETA
                myChar.Text = ChrW(&H398)
            Case &H52                            ' # GREEK CAPITAL LETTER RHO
                myChar.Text = ChrW(&H3A1)
            Case &H53                            ' # GREEK CAPITAL LETTER SIGMA
                myChar.Text = ChrW(&H3A3)
            Case &H54                            ' # GREEK CAPITAL LETTER TAU
                myChar.Text = ChrW(&H3A4)
            Case &H55                            ' # GREEK CAPITAL LETTER UPSILON
                myChar.Text = ChrW(&H3A5)
            Case &H56                            ' # GREEK SMALL LETTER FINAL SIGMA
                myChar.Text = ChrW(&H3C2)
            Case &H57                            ' # GREEK CAPITAL LETTER OMEGA
                myChar.Text = ChrW(&H3A9)
            Case &H57                            ' # OHM SIGN
                myChar.Text = ChrW(&H2126)
            Case &H58                            ' # GREEK CAPITAL LETTER XI
                myChar.Text = ChrW(&H39E)
            Case &H59                            ' # GREEK CAPITAL LETTER PSI
                myChar.Text = ChrW(&H3A8)
            Case &H5A                            ' # GREEK CAPITAL LETTER ZETA
                myChar.Text = ChrW(&H396)
            Case &H5C                            ' # THEREFORE
                myChar.Text = ChrW(&H2234)
            Case &H5E                            ' # UP TACK
                myChar.Text = ChrW(&H22A5)
            Case &H60                            ' # RADICAL EXTENDER
                myChar.Text = ChrW(&HF8E5)
            Case &H61                            ' # GREEK SMALL LETTER ALPHA
                myChar.Text = ChrW(&H3B1)
            Case &H62                            ' # GREEK SMALL LETTER BETA
                myChar.Text = ChrW(&H3B2)
            Case &H63                            ' # GREEK SMALL LETTER CHI
                myChar.Text = ChrW(&H3C7)
            Case &H64                            ' # GREEK SMALL LETTER DELTA
                myChar.Text = ChrW(&H3B4)
            Case &H65                            ' # GREEK SMALL LETTER EPSILON
                myChar.Text = ChrW(&H3B5)
            Case &H66                            ' # GREEK SMALL LETTER PHI
                myChar.Text = ChrW(&H3C6)
            Case &H67                            ' # GREEK SMALL LETTER GAMMA
                myChar.Text = ChrW(&H3B3)
            Case &H68                            ' # GREEK SMALL LETTER ETA
                myChar.Text = ChrW(&H3B7)
            Case &H69                            ' # GREEK SMALL LETTER IOTA
                myChar.Text = ChrW(&H3B9)
            Case &H6A                            ' # GREEK PHI SYMBOL
                myChar.Text = ChrW(&H3D5)
            Case &H6B                            ' # GREEK SMALL LETTER KAPPA
                myChar.Text = ChrW(&H3BA)
            Case &H6C                            ' # GREEK SMALL LETTER LAMDA
                myChar.Text = ChrW(&H3BB)
            Case &H6D                            ' # MICRO SIGN
                myChar.Text = ChrW(&HB5)
            Case &H6D                            ' # GREEK SMALL LETTER MU
                myChar.Text = ChrW(&H3BC)
            Case &H6E                            ' # GREEK SMALL LETTER NU
                myChar.Text = ChrW(&H3BD)
            Case &H6F                            ' # GREEK SMALL LETTER OMICRON
                myChar.Text = ChrW(&H3BF)
            Case &H70                            ' # GREEK SMALL LETTER PI
                myChar.Text = ChrW(&H3C0)
            Case &H71                            ' # GREEK SMALL LETTER THETA
                myChar.Text = ChrW(&H3B8)
            Case &H72                            ' # GREEK SMALL LETTER RHO
                myChar.Text = ChrW(&H3C1)
            Case &H73                            ' # GREEK SMALL LETTER SIGMA
                myChar.Text = ChrW(&H3C3)
            Case &H74                            ' # GREEK SMALL LETTER TAU
                myChar.Text = ChrW(&H3C4)
            Case &H75                            ' # GREEK SMALL LETTER UPSILON
                myChar.Text = ChrW(&H3C5)
            Case &H76                            ' # GREEK PI SYMBOL
                myChar.Text = ChrW(&H3D6)
            Case &H77                            ' # GREEK SMALL LETTER OMEGA
                myChar.Text = ChrW(&H3C9)
            Case &H78                            ' # GREEK SMALL LETTER XI
                myChar.Text = ChrW(&H3BE)
            Case &H79                            ' # GREEK SMALL LETTER PSI
                myChar.Text = ChrW(&H3C8)
            Case &H7A                            ' # GREEK SMALL LETTER ZETA
                myChar.Text = ChrW(&H3B6)
            Case &H7E                            ' # TILDE OPERATOR
                myChar.Text = ChrW(&H223C)
            Case &HA0                            ' # EURO SIGN
                myChar.Text = ChrW(&H20AC)
            Case &HA1                            ' # GREEK UPSILON WITH HOOK SYMBOL
                myChar.Text = ChrW(&H3D2)
            Case &HA2                            ' # PRIME
                myChar.Text = ChrW(&H2032)
            Case &HA3                            ' # LESS-THAN OR EQUAL TO
                myChar.Text = ChrW(&H2264)
            Case &HA4                            ' # FRACTION SLASH
                myChar.Text = ChrW(&H2044)
            Case &HA4                            ' # DIVISION SLASH
                myChar.Text = ChrW(&H2215)
            Case &HA5                            ' # INFINITY
                myChar.Text = ChrW(&H221E)
            Case &HA6                            ' # LATIN SMALL LETTER F WITH HOOK
                myChar.Text = ChrW(&H192)
            Case &HA7                            ' # BLACK CLUB SUIT
                myChar.Text = ChrW(&H2663)
            Case &HA8                            ' # BLACK DIAMOND SUIT
                myChar.Text = ChrW(&H2666)
            Case &HA9                            ' # BLACK HEART SUIT
                myChar.Text = ChrW(&H2665)
            Case &HAA                            ' # BLACK SPADE SUIT
                myChar.Text = ChrW(&H2660)
            Case &HAB                            ' # LEFT RIGHT ARROW
                myChar.Text = ChrW(&H2194)
            Case &HAC                            ' # LEFTWARDS ARROW
                myChar.Text = ChrW(&H2190)
            Case &HAD                            ' # UPWARDS ARROW
                myChar.Text = ChrW(&H2191)
            Case &HAE                            ' # RIGHTWARDS ARROW
                myChar.Text = ChrW(&H2192)
            Case &HAF                            ' # DOWNWARDS ARROW
                myChar.Text = ChrW(&H2193)
            Case &HB2                            ' # DOUBLE PRIME
                myChar.Text = ChrW(&H2033)
            Case &HB3                            ' # GREATER-THAN OR EQUAL TO
                myChar.Text = ChrW(&H2265)
            Case &HB4                            ' # MULTIPLICATION SIGN
                myChar.Text = ChrW(&HD7)
            Case &HB5                            ' # PROPORTIONAL TO
                myChar.Text = ChrW(&H221D)
            Case &HB6                            ' # PARTIAL DIFFERENTIAL
                myChar.Text = ChrW(&H2202)
            Case &HB7                            ' # BULLET
                myChar.Text = ChrW(&H2022)
            Case &HB8                            ' # DIVISION SIGN
                myChar.Text = ChrW(&HF7)
            Case &HB9                            ' # NOT EQUAL TO
                myChar.Text = ChrW(&H2260)
            Case &HBA                            ' # IDENTICAL TO
                myChar.Text = ChrW(&H2261)
            Case &HBB                            ' # ALMOST EQUAL TO
                myChar.Text = ChrW(&H2248)
            Case &HBC                            ' # HORIZONTAL ELLIPSIS
                myChar.Text = ChrW(&H2026)
            Case &HBD                            ' # VERTICAL ARROW EXTENDER
                myChar.Text = ChrW(&HF8E6)
            Case &HBE                            ' # HORIZONTAL ARROW EXTENDER
                myChar.Text = ChrW(&HF8E7)
            Case &HBF                            ' # DOWNWARDS ARROW WITH CORNER LEFTWARDS
                myChar.Text = ChrW(&H21B5)
            Case &HC0                            ' # ALEF SYMBOL
                myChar.Text = ChrW(&H2135)
            Case &HC1                            ' # BLACK-LETTER CAPITAL I
                myChar.Text = ChrW(&H2111)
            Case &HC2                            ' # BLACK-LETTER CAPITAL R
                myChar.Text = ChrW(&H211C)
            Case &HC3                            ' # SCRIPT CAPITAL P
                myChar.Text = ChrW(&H2118)
            Case &HC4                            ' # CIRCLED TIMES
                myChar.Text = ChrW(&H2297)
            Case &HC5                            ' # CIRCLED PLUS
                myChar.Text = ChrW(&H2295)
            Case &HC6                            ' # EMPTY SET
                myChar.Text = ChrW(&H2205)
            Case &HC7                            ' # INTERSECTION
                myChar.Text = ChrW(&H2229)
            Case &HC8                            ' # UNION
                myChar.Text = ChrW(&H222A)
            Case &HC9                            ' # SUPERSET OF
                myChar.Text = ChrW(&H2283)
            Case &HCA                            ' # SUPERSET OF OR EQUAL TO
                myChar.Text = ChrW(&H2287)
            Case &HCB                            ' # NOT A SUBSET OF
                myChar.Text = ChrW(&H2284)
            Case &HCC                            ' # SUBSET OF
                myChar.Text = ChrW(&H2282)
            Case &HCD                            ' # SUBSET OF OR EQUAL TO
                myChar.Text = ChrW(&H2286)
            Case &HCE                            ' # ELEMENT OF
                myChar.Text = ChrW(&H2208)
            Case &HCF                            ' # NOT AN ELEMENT OF
                myChar.Text = ChrW(&H2209)
            Case &HD0                            ' # ANGLE
                myChar.Text = ChrW(&H2220)
            Case &HD1                            ' # NABLA
                myChar.Text = ChrW(&H2207)
            Case &HD2                            ' # REGISTERED SIGN SERIF
                myChar.Text = ChrW(&HF6DA)
            Case &HD3                            ' # COPYRIGHT SIGN SERIF
                myChar.Text = ChrW(&HF6D9)
            Case &HD4                            ' # TRADE MARK SIGN SERIF
                myChar.Text = ChrW(&HF6DB)
            Case &HD5                            ' # N-ARY PRODUCT
                myChar.Text = ChrW(&H220F)
            Case &HD6                            ' # SQUARE ROOT
                myChar.Text = ChrW(&H221A)
            Case &HD7                            ' # DOT OPERATOR
                myChar.Text = ChrW(&H22C5)
            Case &HD8                            ' # NOT SIGN
                myChar.Text = ChrW(&HAC)
            Case &HD9                            ' # LOGICAL AND
                myChar.Text = ChrW(&H2227)
            Case &HDA                            ' # LOGICAL OR
                myChar.Text = ChrW(&H2228)
            Case &HDB                            ' # LEFT RIGHT DOUBLE ARROW
                myChar.Text = ChrW(&H21D4)
            Case &HDC                            ' # LEFTWARDS DOUBLE ARROW
                myChar.Text = ChrW(&H21D0)
            Case &HDD                            ' # UPWARDS DOUBLE ARROW
                myChar.Text = ChrW(&H21D1)
            Case &HDE                            ' # RIGHTWARDS DOUBLE ARROW
                myChar.Text = ChrW(&H21D2)
            Case &HDF                            ' # DOWNWARDS DOUBLE ARROW
                myChar.Text = ChrW(&H21D3)
            Case &HE0                            ' # LOZENGE
                myChar.Text = ChrW(&H25CA)
            Case &HE1                            ' # LEFT-POINTING ANGLE BRACKET
                myChar.Text = ChrW(&H2329)
            Case &HE2                            ' # REGISTERED SIGN SANS SERIF
                myChar.Text = ChrW(&HF8E8)
            Case &HE3                            ' # COPYRIGHT SIGN SANS SERIF
                myChar.Text = ChrW(&HF8E9)
            Case &HE4                            ' # TRADE MARK SIGN SANS SERIF
                myChar.Text = ChrW(&HF8EA)
            Case &HE5                            ' # N-ARY SUMMATION
                myChar.Text = ChrW(&H2211)
            Case &HE6                            ' # LEFT PAREN TOP
                myChar.Text = ChrW(&HF8EB)
            Case &HE7                            ' # LEFT PAREN EXTENDER
                myChar.Text = ChrW(&HF8EC)
            Case &HE8                            ' # LEFT PAREN BOTTOM
                myChar.Text = ChrW(&HF8ED)
            Case &HE9                            ' # LEFT SQUARE BRACKET TOP
                myChar.Text = ChrW(&HF8EE)
            Case &HEA                            ' # LEFT SQUARE BRACKET EXTENDER
                myChar.Text = ChrW(&HF8EF)
            Case &HEB                            ' # LEFT SQUARE BRACKET BOTTOM
                myChar.Text = ChrW(&HF8F0)
            Case &HEC                            ' # LEFT CURLY BRACKET TOP
                myChar.Text = ChrW(&HF8F1)
            Case &HED                            ' # LEFT CURLY BRACKET MID
                myChar.Text = ChrW(&HF8F2)
            Case &HEE                            ' # LEFT CURLY BRACKET BOTTOM
                myChar.Text = ChrW(&HF8F3)
            Case &HEF                            ' # CURLY BRACKET EXTENDER
                myChar.Text = ChrW(&HF8F4)
            Case &HF1                            ' # RIGHT-POINTING ANGLE BRACKET
                myChar.Text = ChrW(&H232A)
            Case &HF2                            ' # INTEGRAL
                myChar.Text = ChrW(&H222B)
            Case &HF3                            ' # TOP HALF INTEGRAL
                myChar.Text = ChrW(&H2320)
            Case &HF4                            ' # INTEGRAL EXTENDER
                myChar.Text = ChrW(&HF8F5)
            Case &HF5                            ' # BOTTOM HALF INTEGRAL
                myChar.Text = ChrW(&H2321)
            Case &HF6                            ' # RIGHT PAREN TOP
                myChar.Text = ChrW(&HF8F6)
            Case &HF7                            ' # RIGHT PAREN EXTENDER
                myChar.Text = ChrW(&HF8F7)
            Case &HF8                            ' # RIGHT PAREN BOTTOM
                myChar.Text = ChrW(&HF8F8)
            Case &HF9                            ' # RIGHT SQUARE BRACKET TOP
                myChar.Text = ChrW(&HF8F9)
            Case &HFA                            ' # RIGHT SQUARE BRACKET EXTENDER
                myChar.Text = ChrW(&HF8FA)
            Case &HFB                            ' # RIGHT SQUARE BRACKET BOTTOM
                myChar.Text = ChrW(&HF8FB)
            Case &HFC                            ' # RIGHT CURLY BRACKET TOP
                myChar.Text = ChrW(&HF8FC)
            Case &HFD                            ' # RIGHT CURLY BRACKET MID
                myChar.Text = ChrW(&HF8FD)
            Case &HFE                            ' # RIGHT CURLY BRACKET BOTTOM
                myChar.Text = ChrW(&HF8FE)
            End Select
            i = i - 1
        End If
    Next myChar
    myRange.Select
End Sub

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Sub UnicodeToSymbol()
    ' ver também https://www.stat.auckland.ac.nz/~paul/R/CM/AdobeSym.html
    ' Source https://groups.google.com/forum/#!topic/microsoft.public.word.printingfonts/pPtEuYqrlvQ
    ' Select document or range before running the macro
    Dim myFont As String
    Dim myCharNum As Long
    Dim myRange As Range
    Dim myChar As Range
    Dim i As Long, CharCount As Long
    Set myRange = Selection.Range.Duplicate
    CharCount = myRange.ComputeStatistics(wdStatisticCharacters)
    For Each myChar In myRange.Characters
        i = i + 1
        If Not (myChar.Font.Name = "Symbol") Then
            mycharN = AscW(myChar.Text)
            myCharNum = mycharN And &HFFFF&
            ' Decorative Fonts are mapped to a
            ' "private use" code page starting at &HF000
            ' myCharNum = myCharNum - &HF000&
            If (mycharN > 879 And mycharN < 1024) Or (mycharN > 2200) Or (mycharN = 215) Then
                Select Case myCharNum
                Case &H22                        ' # FOR ALL
                    myChar.Text = ChrW(&H2200)
                Case &H2203                      ' # THERE EXISTS
                    myChar.Text = ChrW(&H24)
                Case &H220B                      ' # CONTAINS AS MEMBER
                    myChar.Text = ChrW(&H27)
                Case &H2217                      ' # ASTERISK OPERATOR
                    myChar.Text = ChrW(&H2A)
                Case &H2212                      ' # MINUS SIGN
                    myChar.Text = ChrW(&H2D)
                Case &H2245                      ' # APPROXIMATELY EQUAL TO
                    myChar.Text = ChrW(&H40)
                Case &H391                       ' # GREEK CAPITAL LETTER ALPHA
                    myChar.Text = ChrW(&H41)
                Case &H392                       ' # GREEK CAPITAL LETTER BETA
                    myChar.Text = ChrW(&H42)
                Case &H3A7                       ' # GREEK CAPITAL LETTER CHI
                    myChar.Text = ChrW(&H43)
                Case &H394                       ' # GREEK CAPITAL LETTER DELTA
                    myChar.Text = ChrW(&H44)
                Case &H2206                      ' # INCREMENT
                    myChar.Text = ChrW(&H44)
                Case &H395                       ' # GREEK CAPITAL LETTER EPSILON
                    myChar.Text = ChrW(&H45)
                Case &H3A6                       ' # GREEK CAPITAL LETTER PHI
                    myChar.Text = ChrW(&H46)
                Case &H393                       ' # GREEK CAPITAL LETTER GAMMA
                    myChar.Text = ChrW(&H47)
                Case &H397                       ' # GREEK CAPITAL LETTER ETA
                    myChar.Text = ChrW(&H48)
                Case &H399                       ' # GREEK CAPITAL LETTER IOTA
                    myChar.Text = ChrW(&H49)
                Case &H3D1                       ' # GREEK THETA SYMBOL
                    myChar.Text = ChrW(&H4A)
                Case &H39A                       ' # GREEK CAPITAL LETTER KAPPA
                    myChar.Text = ChrW(&H4B)
                Case &H39B                       ' # GREEK CAPITAL LETTER LAMDA
                    myChar.Text = ChrW(&H4C)
                Case &H39C                       ' # GREEK CAPITAL LETTER MU
                    myChar.Text = ChrW(&H4D)
                Case &H39D                       ' # GREEK CAPITAL LETTER NU
                    myChar.Text = ChrW(&H4E)
                Case &H39F                       ' # GREEK CAPITAL LETTER OMICRON
                    myChar.Text = ChrW(&H4F)
                Case &H3A0                       ' # GREEK CAPITAL LETTER PI
                    myChar.Text = ChrW(&H50)
                Case &H398                       ' # GREEK CAPITAL LETTER THETA
                    myChar.Text = ChrW(&H51)
                Case &H3A1                       ' # GREEK CAPITAL LETTER RHO
                    myChar.Text = ChrW(&H52)
                Case &H3A3                       ' # GREEK CAPITAL LETTER SIGMA
                    myChar.Text = ChrW(&H53)
                Case &H3A4                       ' # GREEK CAPITAL LETTER TAU
                    myChar.Text = ChrW(&H54)
                Case &H3A5                       ' # GREEK CAPITAL LETTER UPSILON
                    myChar.Text = ChrW(&H55)
                Case &H3C2                       ' # GREEK SMALL LETTER FINAL SIGMA
                    myChar.Text = ChrW(&H56)
                Case &H3A9                       ' # GREEK CAPITAL LETTER OMEGA
                    myChar.Text = ChrW(&H57)
                Case &H2126                      ' # OHM SIGN
                    myChar.Text = ChrW(&H57)
                Case &H39E                       ' # GREEK CAPITAL LETTER XI
                    myChar.Text = ChrW(&H58)
                Case &H3A8                       ' # GREEK CAPITAL LETTER PSI
                    myChar.Text = ChrW(&H59)
                Case &H396                       ' # GREEK CAPITAL LETTER ZETA
                    myChar.Text = ChrW(&H5A)
                Case &H2234                      ' # THEREFORE
                    myChar.Text = ChrW(&H5C)
                Case &H22A5                      ' # UP TACK
                    myChar.Text = ChrW(&H5E)
                Case &HF8E5                      ' # RADICAL EXTENDER
                    myChar.Text = ChrW(&H60)
                Case &H3B1                       ' # GREEK SMALL LETTER ALPHA
                    myChar.Text = ChrW(&H61)
                Case &H3B2                       ' # GREEK SMALL LETTER BETA
                    myChar.Text = ChrW(&H62)
                Case &H3C7                       ' # GREEK SMALL LETTER CHI
                    myChar.Text = ChrW(&H63)
                Case &H3B4                       ' # GREEK SMALL LETTER DELTA
                    myChar.Text = ChrW(&H64)
                Case &H3B5                       ' # GREEK SMALL LETTER EPSILON
                    myChar.Text = ChrW(&H65)
                Case &H3C6                       ' # GREEK SMALL LETTER PHI
                    myChar.Text = ChrW(&H66)
                Case &H3B3                       ' # GREEK SMALL LETTER GAMMA
                    myChar.Text = ChrW(&H67)
                Case &H3B7                       ' # GREEK SMALL LETTER ETA
                    myChar.Text = ChrW(&H68)
                Case &H3B9                       ' # GREEK SMALL LETTER IOTA
                    myChar.Text = ChrW(&H69)
                Case &H3D5                       ' # GREEK PHI SYMBOL
                    myChar.Text = ChrW(&H6A)
                Case &H3BA                       ' # GREEK SMALL LETTER KAPPA
                    myChar.Text = ChrW(&H6B)
                Case &H3BB                       ' # GREEK SMALL LETTER LAMDA
                    myChar.Text = ChrW(&H6C)
                Case &HB5                        ' # MICRO SIGN
                    myChar.Text = ChrW(&H6D)
                Case &H3BC                       ' # GREEK SMALL LETTER MU
                    myChar.Text = ChrW(&H6D)
                Case &H3BD                       ' # GREEK SMALL LETTER NU
                    myChar.Text = ChrW(&H6E)
                Case &H3BF                       ' # GREEK SMALL LETTER OMICRON
                    myChar.Text = ChrW(&H6F)
                Case &H3C0                       ' # GREEK SMALL LETTER PI
                    myChar.Text = ChrW(&H70)
                Case &H3B8                       ' # GREEK SMALL LETTER THETA
                    myChar.Text = ChrW(&H71)
                Case &H3C1                       ' # GREEK SMALL LETTER RHO
                    myChar.Text = ChrW(&H72)
                Case &H3C3                       ' # GREEK SMALL LETTER SIGMA
                    myChar.Text = ChrW(&H73)
                Case &H3C4                       ' # GREEK SMALL LETTER TAU
                    myChar.Text = ChrW(&H74)
                Case &H3C5                       ' # GREEK SMALL LETTER UPSILON
                    myChar.Text = ChrW(&H75)
                Case &H3D6                       ' # GREEK PI SYMBOL
                    myChar.Text = ChrW(&H76)
                Case &H3C9                       ' # GREEK SMALL LETTER OMEGA
                    myChar.Text = ChrW(&H77)
                Case &H3BE                       ' # GREEK SMALL LETTER XI
                    myChar.Text = ChrW(&H78)
                Case &H3C8                       ' # GREEK SMALL LETTER PSI
                    myChar.Text = ChrW(&H79)
                Case &H3B6                       ' # GREEK SMALL LETTER ZETA
                    myChar.Text = ChrW(&H7A)
                Case &H223C                      ' # TILDE OPERATOR
                    myChar.Text = ChrW(&H7E)
                Case &H20AC                      ' # EURO SIGN
                    myChar.Text = ChrW(&HA0)
                Case &H3D2                       ' # GREEK UPSILON WITH HOOK SYMBOL
                    myChar.Text = ChrW(&HA1)
                Case &H2032                      ' # PRIME
                    myChar.Text = ChrW(&HA2)
                Case &H2264                      ' # LESS-THAN OR EQUAL TO
                    myChar.Text = ChrW(&HA3)
                Case &H2044                      ' # FRACTION SLASH
                    myChar.Text = ChrW(&HA4)
                Case &H2215                      ' # DIVISION SLASH
                    myChar.Text = ChrW(&HA4)
                Case &H221E                      ' # INFINITY
                    myChar.Text = ChrW(&HA5)
                Case &H192                       ' # LATIN SMALL LETTER F WITH HOOK
                    myChar.Text = ChrW(&HA6)
                Case &H2663                      ' # BLACK CLUB SUIT
                    myChar.Text = ChrW(&HA7)
                Case &H2666                      ' # BLACK DIAMOND SUIT
                    myChar.Text = ChrW(&HA8)
                Case &H2665                      ' # BLACK HEART SUIT
                    myChar.Text = ChrW(&HA9)
                Case &H2660                      ' # BLACK SPADE SUIT
                    myChar.Text = ChrW(&HAA)
                Case &H2194                      ' # LEFT RIGHT ARROW
                    myChar.Text = ChrW(&HAB)
                Case &H2190                      ' # LEFTWARDS ARROW
                    myChar.Text = ChrW(&HAC)
                Case &H2191                      ' # UPWARDS ARROW
                    myChar.Text = ChrW(&HAD)
                Case &H2192                      ' # RIGHTWARDS ARROW
                    myChar.Text = ChrW(&HAE)
                Case &H2193                      ' # DOWNWARDS ARROW
                    myChar.Text = ChrW(&HAF)
                Case &H2033                      ' # DOUBLE PRIME
                    myChar.Text = ChrW(&HB2)
                Case &H2265                      ' # GREATER-THAN OR EQUAL TO
                    myChar.Text = ChrW(&HB3)
                Case &HD7                        ' # MULTIPLICATION SIGN
                    myChar.Text = ChrW(&HB4)
                Case &H221D                      ' # PROPORTIONAL TO
                    myChar.Text = ChrW(&HB5)
                Case &H2202                      ' # PARTIAL DIFFERENTIAL
                    myChar.Text = ChrW(&HB6)
                Case &H2022                      ' # BULLET
                    myChar.Text = ChrW(&HB7)
                Case &HF7                        ' # DIVISION SIGN
                    myChar.Text = ChrW(&HB8)
                Case &H2260                      ' # NOT EQUAL TO
                    myChar.Text = ChrW(&HB9)
                Case &H2261                      ' # IDENTICAL TO
                    myChar.Text = ChrW(&HBA)
                Case &H2248                      ' # ALMOST EQUAL TO
                    myChar.Text = ChrW(&HBB)
                Case &H2026                      ' # HORIZONTAL ELLIPSIS
                    myChar.Text = ChrW(&HBC)
                Case &HF8E6                      ' # VERTICAL ARROW EXTENDER
                    myChar.Text = ChrW(&HBD)
                Case &HF8E7                      ' # HORIZONTAL ARROW EXTENDER
                    myChar.Text = ChrW(&HBE)
                Case &H21B5                      ' # DOWNWARDS ARROW WITH CORNER LEFTWARDS
                    myChar.Text = ChrW(&HBF)
                Case &H2135                      ' # ALEF SYMBOL
                    myChar.Text = ChrW(&HC0)
                Case &H2111                      ' # BLACK-LETTER CAPITAL I
                    myChar.Text = ChrW(&HC1)
                Case &H211C                      ' # BLACK-LETTER CAPITAL R
                    myChar.Text = ChrW(&HC2)
                Case &H2118                      ' # SCRIPT CAPITAL P
                    myChar.Text = ChrW(&HC3)
                Case &H2297                      ' # CIRCLED TIMES
                    myChar.Text = ChrW(&HC4)
                Case &H2295                      ' # CIRCLED PLUS
                    myChar.Text = ChrW(&HC5)
                Case &H2205                      ' # EMPTY SET
                    myChar.Text = ChrW(&HC6)
                Case &H2229                      ' # INTERSECTION
                    myChar.Text = ChrW(&HC7)
                Case &H222A                      ' # UNION
                    myChar.Text = ChrW(&HC8)
                Case &H2283                      ' # SUPERSET OF
                    myChar.Text = ChrW(&HC9)
                Case &H2287                      ' # SUPERSET OF OR EQUAL TO
                    myChar.Text = ChrW(&HCA)
                Case &H2284                      ' # NOT A SUBSET OF
                    myChar.Text = ChrW(&HCB)
                Case &H2282                      ' # SUBSET OF
                    myChar.Text = ChrW(&HCC)
                Case &H2286                      ' # SUBSET OF OR EQUAL TO
                    myChar.Text = ChrW(&HCD)
                Case &H2208                      ' # ELEMENT OF
                    myChar.Text = ChrW(&HCE)
                Case &H2209                      ' # NOT AN ELEMENT OF
                    myChar.Text = ChrW(&HCF)
                Case &H2220                      ' # ANGLE
                    myChar.Text = ChrW(&HD0)
                Case &H2207                      ' # NABLA
                    myChar.Text = ChrW(&HD1)
                Case &HF6DA                      ' # REGISTERED SIGN SERIF
                    myChar.Text = ChrW(&HD2)
                Case &HF6D9                      ' # COPYRIGHT SIGN SERIF
                    myChar.Text = ChrW(&HD3)
                Case &HF6DB                      ' # TRADE MARK SIGN SERIF
                    myChar.Text = ChrW(&HD4)
                Case &H220F                      ' # N-ARY PRODUCT
                    myChar.Text = ChrW(&HD5)
                Case &H221A                      ' # SQUARE ROOT
                    myChar.Text = ChrW(&HD6)
                Case &H22C5                      ' # DOT OPERATOR
                    myChar.Text = ChrW(&HD7)
                Case &HAC                        ' # NOT SIGN
                    myChar.Text = ChrW(&HD8)
                Case &H2227                      ' # LOGICAL AND
                    myChar.Text = ChrW(&HD9)
                Case &H2228                      ' # LOGICAL OR
                    myChar.Text = ChrW(&HDA)
                Case &H21D4                      ' # LEFT RIGHT DOUBLE ARROW
                    myChar.Text = ChrW(&HDB)
                Case &H21D0                      ' # LEFTWARDS DOUBLE ARROW
                    myChar.Text = ChrW(&HDC)
                Case &H21D1                      ' # UPWARDS DOUBLE ARROW
                    myChar.Text = ChrW(&HDD)
                Case &H21D2                      ' # RIGHTWARDS DOUBLE ARROW
                    myChar.Text = ChrW(&HDE)
                Case &H21D3                      ' # DOWNWARDS DOUBLE ARROW
                    myChar.Text = ChrW(&HDF)
                Case &H25CA                      ' # LOZENGE
                    myChar.Text = ChrW(&HE0)
                Case &H2329                      ' # LEFT-POINTING ANGLE BRACKET
                    myChar.Text = ChrW(&HE1)
                Case &HF8E8                      ' # REGISTERED SIGN SANS SERIF
                    myChar.Text = ChrW(&HE2)
                Case &HF8E9                      ' # COPYRIGHT SIGN SANS SERIF
                    myChar.Text = ChrW(&HE3)
                Case &HF8EA                      ' # TRADE MARK SIGN SANS SERIF
                    myChar.Text = ChrW(&HE4)
                Case &H2211                      ' # N-ARY SUMMATION
                    myChar.Text = ChrW(&HE5)
                Case &HF8EB                      ' # LEFT PAREN TOP
                    myChar.Text = ChrW(&HE6)
                Case &HF8EC                      ' # LEFT PAREN EXTENDER
                    myChar.Text = ChrW(&HE7)
                Case &HF8ED                      ' # LEFT PAREN BOTTOM
                    myChar.Text = ChrW(&HE8)
                Case &HF8EE                      ' # LEFT SQUARE BRACKET TOP
                    myChar.Text = ChrW(&HE9)
                Case &HF8EF                      ' # LEFT SQUARE BRACKET EXTENDER
                    myChar.Text = ChrW(&HEA)
                Case &HF8F0                      ' # LEFT SQUARE BRACKET BOTTOM
                    myChar.Text = ChrW(&HEB)
                Case &HF8F1                      ' # LEFT CURLY BRACKET TOP
                    myChar.Text = ChrW(&HEC)
                Case &HF8F2                      ' # LEFT CURLY BRACKET MID
                    myChar.Text = ChrW(&HED)
                Case &HF8F3                      ' # LEFT CURLY BRACKET BOTTOM
                    myChar.Text = ChrW(&HEE)
                Case &HF8F4                      ' # CURLY BRACKET EXTENDER
                    myChar.Text = ChrW(&HEF)
                Case &H232A                      ' # RIGHT-POINTING ANGLE BRACKET
                    myChar.Text = ChrW(&HF1)
                Case &H222B                      ' # INTEGRAL
                    myChar.Text = ChrW(&HF2)
                Case &H2320                      ' # TOP HALF INTEGRAL
                    myChar.Text = ChrW(&HF3)
                Case &HF8F5                      ' # INTEGRAL EXTENDER
                    myChar.Text = ChrW(&HF4)
                Case &H2321                      ' # BOTTOM HALF INTEGRAL
                    myChar.Text = ChrW(&HF5)
                Case &HF8F6                      ' # RIGHT PAREN TOP
                    myChar.Text = ChrW(&HF6)
                Case &HF8F7                      ' # RIGHT PAREN EXTENDER
                    myChar.Text = ChrW(&HF7)
                Case &HF8F8                      ' # RIGHT PAREN BOTTOM
                    myChar.Text = ChrW(&HF8)
                Case &HF8F9                      ' # RIGHT SQUARE BRACKET TOP
                    myChar.Text = ChrW(&HF9)
                Case &HF8FA                      ' # RIGHT SQUARE BRACKET EXTENDER
                    myChar.Text = ChrW(&HFA)
                Case &HF8FB                      ' # RIGHT SQUARE BRACKET BOTTOM
                    myChar.Text = ChrW(&HFB)
                Case &HF8FC                      ' # RIGHT CURLY BRACKET TOP
                    myChar.Text = ChrW(&HFC)
                Case &HF8FD                      ' # RIGHT CURLY BRACKET MID
                    myChar.Text = ChrW(&HFD)
                Case &HF8FE                      ' # RIGHT CURLY BRACKET BOTTOM
                    myChar.Text = ChrW(&HFE)
                Case &H3016                      ' # left parentheses like cumbria math
                    myChar.Text = "("
                Case &H3017                      ' # right parentheses like cumbria math
                    myChar.Text = ")"
                End Select
                myChar.Font.Name = "Symbol"
            End If
            i = i - 1
        End If
    Next myChar
    myRange.Select
End Sub

Sub ordinal()
    '
    ' ordinal Macro
    ' substitui letra o sublinha superscrita pelo símbolo de ordinal º
    '
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Underline = wdUnderlineSingle
        .Superscript = True
        .Subscript = False
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Underline = wdUnderlineNone
        .Superscript = False
        .Subscript = False
    End With
    With Selection.Find
        .Text = "o"
        .Replacement.Text = "º"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "a"
        .Replacement.Text = "ª"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub tabs()
    '
    ' tabs Macro
    ' substitui tabs por espaços
    '
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "  "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub indent()
    ' formata  recuo máximo
    Dim sDefaultIndent As Single
    Dim sCalcul As Single
    Dim opara As Paragraph
    sDefaultIndent = CentimetersToPoints(1.5)
    For Each opara In ActiveDocument.Paragraphs
        If opara.FirstLineIndent > sDefaultIndent Then
            opara.FirstLineIndent = sDefaultIndent
        End If
    Next
End Sub


Sub formata_dou()
    ' executa sub em sequência
    ' importante: acione os controles de revisão para verificar o que foi feito
    ' TODO: mudar curly quotes por simple quotes
    Application.ScreenUpdating = False
    calibri9
    converter_tabela_12
    indent
    tabs
    ordinal
    converte_rodape
    Application.ScreenUpdating = True
    'Todo incluir message box
End Sub
