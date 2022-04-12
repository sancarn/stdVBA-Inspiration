Attribute VB_Name = "modFuzzy"
Option Explicit
Type RankInfo
    Offset As Integer
    Percentage As Single
End Type

Function FuzzyPercent(ByVal String1 As String, _
                      ByVal String2 As String, _
                      Optional Algorithm As Integer = 3, _
                      Optional Normalised As Boolean = False) As Single
    '*************************************
    '** Return a % match on two strings **
    '*************************************
    Dim intLen1 As Integer, intLen2 As Integer
    Dim intCurLen As Integer
    Dim intTo As Integer
    Dim intPos As Integer
    Dim intPtr As Integer
    Dim intScore As Integer
    Dim intTotScore As Integer
    Dim intStartPos As Integer
    Dim strWork As String

    '-------------------------------------------------------
    '-- If strings havent been normalised, normalise them --
    '-------------------------------------------------------
    If Normalised = False Then
        String1 = LCase(Trim(String1))
        String2 = LCase(Trim(String2))
    End If

    '----------------------------------------------
    '-- Give 100% match if strings exactly equal --
    '----------------------------------------------
    If String1 = String2 Then
        FuzzyPercent = 1
        Exit Function
    End If

    intLen1 = Len(String1)
    intLen2 = Len(String2)

    '----------------------------------------
    '-- Give 0% match if string length < 2 --
    '----------------------------------------
    If intLen1 < 2 Then
        FuzzyPercent = 0
        Exit Function
    End If

    intTotScore = 0                              'initialise total possible score
    intScore = 0                                 'initialise current score

    '--------------------------------------------------------
    '-- If Algorithm = 1 or 3, Search for single characters --
    '--------------------------------------------------------
    If (Algorithm And 1) <> 0 Then
        FuzzyAlg1 String1, String2, intScore, intTotScore
        If intLen1 < intLen2 Then FuzzyAlg1 String2, String1, intScore, intTotScore
    End If

    '-----------------------------------------------------------
    '-- If Algorithm = 2 or 3, Search for pairs, triplets etc. --
    '-----------------------------------------------------------
    If (Algorithm And 2) <> 0 Then
        FuzzyAlg2 String1, String2, intScore, intTotScore
        If intLen1 < intLen2 Then FuzzyAlg2 String2, String1, intScore, intTotScore
    End If

    FuzzyPercent = intScore / intTotScore

End Function

Private Sub FuzzyAlg1(ByVal String1 As String, _
                      ByVal String2 As String, _
                      ByRef score As Integer, _
                      ByRef TotScore As Integer)
    Dim intLen1 As Integer, intPos As Integer, intPtr As Integer, intStartPos As Integer

    intLen1 = Len(String1)
    TotScore = TotScore + intLen1                'update total possible score
    intPos = 0
    For intPtr = 1 To intLen1
        intStartPos = intPos + 1
        intPos = InStr(intStartPos, String2, Mid$(String1, intPtr, 1))
        If intPos > 0 Then
            If intPos > intStartPos + 3 Then     'No match if char is > 3 bytes away
                intPos = intStartPos
            Else
                score = score + 1                'Update current score
            End If
        Else
            intPos = intStartPos
        End If
    Next intPtr
End Sub

Private Sub FuzzyAlg2(ByVal String1 As String, _
                      ByVal String2 As String, _
                      ByRef score As Integer, _
                      ByRef TotScore As Integer)
    Dim intCurLen As Integer, intLen1 As Integer, intTo As Integer, intPtr As Integer, intPos As Integer
    Dim strWork As String

    intLen1 = Len(String1)
    For intCurLen = 2 To intLen1
        strWork = String2                        'Get a copy of String2
        intTo = intLen1 - intCurLen + 1
        TotScore = TotScore + Int(intLen1 / intCurLen) 'Update total possible score
        For intPtr = 1 To intTo Step intCurLen
            intPos = InStr(strWork, Mid$(String1, intPtr, intCurLen))
            If intPos > 0 Then
                Mid$(strWork, intPos, intCurLen) = String$(intCurLen, &H0) 'corrupt found string
                score = score + 1                'Update current score
            End If
        Next intPtr
    Next intCurLen

End Sub

Function FuzzyVLookup(ByVal LookupValue As String, _
                      ByVal TableArray As Range, _
                      ByVal IndexNum As Integer, _
                      Optional NFPercent As Single = 0.05, _
                      Optional Rank As Integer = 1, _
                      Optional Algorithm As Integer = 3, _
                      Optional AdditionalCols As Integer = 0) As Variant
    '********************************************************************************
    '** Function to Fuzzy match LookupValue with entries in                        **
    '** column 1 of table specified by TableArray.                                 **
    '** TableArray must specify the top left cell of the range to be searched      **
    '** The function stops scanning the table when an empty cell in column 1       **
    '** is found.                                                                  **
    '** For each entry in column 1 of the table, FuzzyPercent is called to match   **
    '** LookupValue with the Table entry.                                          **
    '** 'Rank' is an optional parameter which may take any value > 0               **
    '**        (default 1) and causes the function to return the 'nth' best        **
    '**         match (where 'n' is defined by 'Rank' parameter)                   **
    '** If the 'Rank' match percentage < NFPercent (Default 5%), #N/A is returned. **
    '** IndexNum is the column number of the entry in TableArray required to be    **
    '** returned, as follows:                                                      **
    '** If IndexNum > 0 and the 'Rank' percentage match is >= NFPercent            **
    '**                 (Default 5%) the column entry indicated by IndexNum is     **
    '**                 returned.                                                  **
    '** if IndexNum = 0 and the 'Rank' percentage match is >= NFPercent            **
    '**                 (Default 5%) the offset row (starting at 1) is returned.   **
    '**                 This value can be used directly in the 'Index' function.   **
    '**                                                                            **
    '** Algorithm can take one of the following values:                            **
    '** Algorithm = 1:                                                             **
    '**     This algorithm is best suited for matching mis-spellings.              **
    '**     For each character in 'String1', a search is performed on 'String2'.   **
    '**     The search is deemed successful if a character is found in 'String2'   **
    '**     within 3 characters of the current position.                           **
    '**     A score is kept of matching characters which is returned as a          **
    '**     percentage of the total possible score.                                **
    '** Algorithm = 2:                                                             **
    '**     This algorithm is best suited for matching sentences, or               **
    '**     'firstname lastname' compared with 'lastname firstname' combinations   **
    '**     A count of matching pairs, triplets, quadruplets etc. in 'String1' and **
    '**     'String2' is returned as a percentage of the total possible.           **
    '** Algorithm = 3: Both Algorithms 1 and 2 are performed.                      **
    '********************************************************************************
    Dim R As Range

    Dim strListString As String
    Dim strWork As String

    Dim sngMinPercent As Single
    Dim sngWork As Single
    Dim sngCurPercent  As Single
    Dim intBestMatchPtr As Integer
    Dim intRankPtr As Integer
    Dim intRankPtr1 As Integer
    Dim i As Integer

    Dim lEndRow As Long

    Dim udRankData() As RankInfo

    Dim vCurValue As Variant

    '--------------------------------------------------------------
    '--    Validation                                            --
    '--------------------------------------------------------------

    LookupValue = LCase$(Application.Trim(LookupValue))

    If IsMissing(NFPercent) Then
        sngMinPercent = 0.05
    Else
        If (NFPercent <= 0) Or (NFPercent > 1) Then
            FuzzyVLookup = "*** 'NFPercent' must be a percentage > zero ***"
            Exit Function
        End If
        sngMinPercent = NFPercent
    End If

    If Rank < 1 Then
        FuzzyVLookup = "*** 'Rank' must be an integer > 0 ***"
        Exit Function
    End If

    ReDim udRankData(1 To Rank)

    lEndRow = TableArray.Rows.Count
    If VarType(TableArray.Cells(lEndRow, 1).Value) = vbEmpty Then
        lEndRow = TableArray.Cells(lEndRow, 1).End(xlUp).Row
    End If

    '---------------
    '-- Main loop --
    '---------------
    For Each R In Range(TableArray.Cells(1, 1), TableArray.Cells(lEndRow, 1))
        vCurValue = ""
        For i = 0 To AdditionalCols
            vCurValue = vCurValue & R.Offset(0, i).Text
        Next i
        If VarType(vCurValue) = vbString Then
            strListString = LCase$(Application.Trim(vCurValue))
        
            '------------------------------------------------
            '-- Fuzzy match strings & get percentage match --
            '------------------------------------------------
            sngCurPercent = FuzzyPercent(String1:=LookupValue, _
                                         String2:=strListString, _
                                         Algorithm:=Algorithm, _
                                         Normalised:=True)
        
            If sngCurPercent >= sngMinPercent Then
                '---------------------------
                '-- Store in ranked array --
                '---------------------------
                For intRankPtr = 1 To Rank
                    If sngCurPercent > udRankData(intRankPtr).Percentage Then
                        For intRankPtr1 = Rank To intRankPtr + 1 Step -1
                            With udRankData(intRankPtr1)
                                .Offset = udRankData(intRankPtr1 - 1).Offset
                                .Percentage = udRankData(intRankPtr1 - 1).Percentage
                            End With
                        Next intRankPtr1
                        With udRankData(intRankPtr)
                            .Offset = R.Row
                            .Percentage = sngCurPercent
                        End With
                        Exit For
                    End If
                Next intRankPtr
            End If
        
        End If
    Next R

    If udRankData(Rank).Percentage < sngMinPercent Then
        '--------------------------------------
        '-- Return '#N/A' if below NFPercent --
        '--------------------------------------
        FuzzyVLookup = CVErr(xlErrNA)
    Else
        intBestMatchPtr = udRankData(Rank).Offset - TableArray.Cells(1, 1).Row + 1
        If IndexNum > 0 Then
            '-----------------------------------
            '-- Return column entry specified --
            '-----------------------------------
            FuzzyVLookup = TableArray.Cells(intBestMatchPtr, IndexNum)
        Else
            '-----------------------
            '-- Return offset row --
            '-----------------------
            FuzzyVLookup = intBestMatchPtr
        End If
    End If
End Function

Function FuzzyHLookup(ByVal LookupValue As String, _
                      ByVal TableArray As Range, _
                      ByVal IndexNum As Integer, _
                      Optional NFPercent As Single = 0.05, _
                      Optional Rank As Integer = 1, _
                      Optional Algorithm As Integer = 3) As Variant
    '********************************************************************************
    '** Function to Fuzzy match LookupValue with entries in                        **
    '** row 1 of table specified by TableArray.                                    **
    '** TableArray must specify the top left cell of the range to be searched      **
    '** The function stops scanning the table when an empty cell in row 1          **
    '** is found.                                                                  **
    '** For each entry in row 1 of the table, FuzzyPercent is called to match      **
    '** LookupValue with the Table entry.                                          **
    '** 'Rank' is an optional parameter which may take any value > 0               **
    '**        (default 1) and causes the function to return the 'nth' best        **
    '**         match (where 'n' is defined by 'Rank' parameter)                   **
    '** If the 'Rank' match percentage < NFPercent (Default 5%), #N/A is returned. **
    '** IndexNum is the row number of the entry in TableArray required to be       **
    '** returned, as follows:                                                      **
    '** If IndexNum > 0 and the 'Rank' percentage match is >= NFPercent            **
    '**                 (Default 5%) the row entry indicated by IndexNum is        **
    '**                 returned.                                                  **
    '** if IndexNum = 0 and the 'Rank' percentage match is >= NFPercent            **
    '**                 (Default 5%) the offset col (starting at 0) is returned.   **
    '**                 This value can be used directly in the 'OffSet' function.  **
    '**                                                                            **
    '** Algorithm can take one of the following values:                            **
    '** Algorithm = 1:                                                             **
    '**     For each character in 'String1', a search is performed on 'String2'.   **
    '**     The search is deemed successful if a character is found in 'String2'   **
    '**     within 3 characters of the current position.                           **
    '**     A score is kept of matching characters which is returned as a          **
    '**     percentage of the total possible score.                                **
    '** Algorithm = 2:                                                             **
    '**     A count of matching pairs, triplets, quadruplets etc. in 'String1' and **
    '**     'String2' is returned as a percentage of the total possible.           **
    '** Algorithm = 3: Both Algorithms 1 and 2 are performed.                      **
    '********************************************************************************
    Dim R As Range

    Dim strListString As String
    Dim strWork As String

    Dim sngMinPercent As Single
    Dim sngWork As Single
    Dim sngCurPercent  As Single

    Dim intBestMatchPtr As Integer
    Dim intPtr As Integer
    Dim intRankPtr As Integer
    Dim intRankPtr1 As Integer

    Dim iEndCol As Integer

    Dim udRankData() As RankInfo

    Dim vCurValue As Variant
    '--------------------------------------------------------------
    '--    Validation                                            --
    '--------------------------------------------------------------
    LookupValue = LCase$(Application.Trim(LookupValue))

    If IsMissing(NFPercent) Then
        sngMinPercent = 0.05
    Else
        If (NFPercent <= 0) Or (NFPercent > 1) Then
            FuzzyHLookup = "*** 'NFPercent' must be a percentage > zero ***"
            Exit Function
        End If
        sngMinPercent = NFPercent
    End If

    If Rank < 1 Then
        FuzzyHLookup = "*** 'Rank' must be an integer > 0 ***"
        Exit Function
    End If

    ReDim udRankData(1 To Rank)
    '**************************
    iEndCol = TableArray.Columns.Count
    If VarType(TableArray.Cells(1, iEndCol).Value) = vbEmpty Then
        iEndCol = TableArray.Cells(1, iEndCol).End(xlToLeft).Column
    End If

    '---------------
    '-- Main loop --
    '---------------
    For Each R In Range(TableArray.Cells(1, 1), TableArray.Cells(1, iEndCol))
        vCurValue = R.Value
        If VarType(vCurValue) = vbString Then
            strListString = LCase$(Application.Trim(vCurValue))
        
            '------------------------------------------------
            '-- Fuzzy match strings & get percentage match --
            '------------------------------------------------
            sngCurPercent = FuzzyPercent(String1:=LookupValue, _
                                         String2:=strListString, _
                                         Algorithm:=Algorithm, _
                                         Normalised:=True)
        
            If sngCurPercent >= sngMinPercent Then
                '---------------------------
                '-- Store in ranked array --
                '---------------------------
                For intRankPtr = 1 To Rank
                    If sngCurPercent > udRankData(intRankPtr).Percentage Then
                        For intRankPtr1 = Rank To intRankPtr + 1 Step -1
                            With udRankData(intRankPtr1)
                                .Offset = udRankData(intRankPtr1 - 1).Offset
                                .Percentage = udRankData(intRankPtr1 - 1).Percentage
                            End With
                        Next intRankPtr1
                        With udRankData(intRankPtr)
                            .Offset = R.Column
                            .Percentage = sngCurPercent
                        End With
                        Exit For
                    End If
                Next intRankPtr
            End If
        
        End If
    Next R

    If udRankData(Rank).Percentage < sngMinPercent Then
        '--------------------------------------
        '-- Return '#N/A' if below NFPercent --
        '--------------------------------------
        FuzzyHLookup = CVErr(xlErrNA)
    Else
        intBestMatchPtr = udRankData(Rank).Offset - TableArray.Cells(1, 1).Column + 1
        If IndexNum > 0 Then
            '-----------------------------------
            '-- Return row entry specified --
            '-----------------------------------
            FuzzyHLookup = TableArray.Cells(IndexNum, intBestMatchPtr)
        Else
            '-----------------------
            '-- Return offset col --
            '-----------------------
            FuzzyHLookup = intBestMatchPtr
        End If
    End If
End Function

Public Function Levenshtein(S1 As String, S2 As String)

    Dim i As Integer
    Dim j As Integer
    Dim L1 As Integer
    Dim L2 As Integer
    Dim d() As Integer
    Dim min1 As Integer
    Dim min2 As Integer

    L1 = Len(S1)
    L2 = Len(S2)
    ReDim d(L1, L2)
    For i = 0 To L1
        d(i, 0) = i
    Next
    For j = 0 To L2
        d(0, j) = j
    Next
    For i = 1 To L1
        For j = 1 To L2
            If Mid(S1, i, 1) = Mid(S2, j, 1) Then
                d(i, j) = d(i - 1, j - 1)
            Else
                min1 = d(i - 1, j) + 1
                min2 = d(i, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                min2 = d(i - 1, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                d(i, j) = min1
            End If
        Next
    Next
    Levenshtein = d(L1, L2)
End Function


