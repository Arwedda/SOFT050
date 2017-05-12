    Sub CreateCombinations()
        sSixCardCombos = Array("01234", "01235", "01245", "01345", "02345", 12345)
        sSevCardCombos = Array("01234", "01235", "01245", "01345", "02345", 12345, "01236","01246", "01256", "01346", _
        "01356", "01456", "02346", "02356", "02456", "03456", 12346, 12356, 12456, 13456, 23456)
    End Sub

    Sub Init(StartChips)
        iHandNum = 0
        iPot = 0
        iPChips = StartChips
        iC1Chips = StartChips
    End Sub

    Sub Interface()
        Dim HMid
        Dim WMid
        Dim i
        HMid = (document.body.clientHeight - 250) / 2
        WMid = (document.body.clientWidth - 200) / 2
        For i = 0 to 4
            document.getElementById("imgCC" & i).style.posTop = HMid
            document.getElementById("imgCC" & i).style.posLeft = WMid + (210 * (i-2))
        Next
        parResult.style.posTop = HMid

        HMid = (document.body.clientHeight - 250)
        For i = 0 to 1
            document.getElementById("imgPC" & i).style.posTop = HMid
        Next
        imgPC0.style.posLeft = WMid - 105
        imgPC1.style.posLeft = WMid + 105
        parPChips.style.posTop = HMid
        parPChips.style.posLeft = WMid + 310

        For i = 0 to 1
            document.getElementById("imgC1C" & i).style.posTop = 0
        Next
        imgC1C0.style.posLeft = WMid - 105
        imgC1C1.style.posLeft = WMid + 105
        parC1Chips.style.posTop = 0
        parC1Chips.style.posLeft = WMid + 310

        ShowBet
    End Sub
        
    Sub FreshHand()
        iRound = 0
        iPot = 0
        iHandNum = iHandNum + 1
        ClearCards sC1Card, iC1Value, sC1Suit, "imgC1C", 1
        ClearCards sPCard, iPValue, sPSuit, "imgPC", 1
        ClearCards sCCard, iCValue, sCSuit, "imgCC", 4
        ClearScores iC1BestScore
        ClearScores iPBestScore
        ClearScores iCheckScore 
        ClearHandVars
        sHandRes = ""
        sCheckHandDesc = ""
        sC1BestHandDesc = ""
        sPBestHandDesc = ""
        bC1Fold = False
        bC1Moved = False
    End Sub

    Sub ClearCards(Card, Value, Suit, pic, j)
        Dim i
        For i = 0 to j
            Card(i) = ""
            Value(i) = 0
            Suit(i) = ""
            document.getElementById(pic & i).src = "bk.jpg"
        Next
    End Sub

    Sub ClearScores(Scr)
        Dim i
        Scr = 0
        For i = 0 to 4
            iCheckVals(i) = 0
            sCheckSuits(i) = ""
        Next
    End Sub

    Sub ClearHandVars()
        sHighcard = 0
        iPair(0) = 0
        iPair(1) = 0
        iThreeOfAKind = 0
        bStraight = False
        bFlush = False
    End Sub

    Sub DealHCs()
        Do While sPCard(0) = sPCard(1)
            PickCards sPCard, iPValue, sPSuit, 0, 1
        Loop
        Cards.RemoveCard sPCard, 0, 1
        Do While sC1Card(0) = sC1Card(1) Or sC1Card(0) = "" Or sC1Card(1) = ""
            PickCards sC1Card, iC1Value, sC1Suit, 0, 1
        Loop    
        Cards.RemoveCard sC1Card, 0, 1
    End Sub

    Sub DealFlop()
        Do While sCCard(0) = sCCard(1) or sCCard(0) = sCCard(2)  or sCCard(1) = sCCard(2) or sCCard(0) = "" or sCCard(1) = "" or sCCard(2) = ""
            PickCards sCCard, iCValue, sCSuit, 0, 2
        Loop
	    Cards.RemoveCard sCCard, 0, 2
    End Sub

    Sub DealTurn()
        Do While sCCard(3) = ""
           PickCards sCCard, iCValue, sCSuit, 3, 3
        Loop
	    Cards.RemoveCard sCCard, 3, 3
    End Sub

    Sub DealRiver()
        Do While sCCard(4) = ""
            PickCards sCCard, iCValue, sCSuit, 4, 4
        Loop
    End Sub

    Sub PickCards(Img, CV, CS, j, k)
        Dim RandNo
        Dim i
        For i = j to k
            RandNo = Int(RND()*52)
            Img(i) = Cards.sPicture(RandNo)
            CV(i) = Cards.iValue(RandNo)
            CS(i) = Cards.sSuit(RandNo)
        Next
    End Sub

    Sub SuitsAndValues(CV, CS, pic, Cards, j, k)
        Dim i
        For i = j to k
            CS(i) = Left(Cards(i), 1)
            CV(i) = CINT(Mid(Cards(i), 2, 2))
        Next
    End Sub

    Sub Sort(CV, CS, Cards, j, ShowPics)
        Dim Swap
        Dim Count
        Dim Temp
        Do
            Swap = False
            For Count = 0 to j
                If CV(Count) > CV(Count+1) Then
                    Swap = True
                    Temp = CV(Count+1)
                    CV(Count+1) = CV(Count)
                    CV(Count) = Temp
                    Temp = CS(Count+1)
                    CS(Count+1) = CS(Count)
                    CS(Count) = Temp
                    If ShowPics = True Then
                        Temp = Cards(Count+1)
                        Cards(Count+1) = Cards(Count)
                        Cards(Count) = Temp
                    End If
                End If
            Next
        Loop Until Swap = False
    End Sub

    Sub ShowCards(pic, Cards, j, k)
        Dim i
            For i = j to k
                document.getElementById(pic & i).src = Cards(i)
            Next
    End Sub

    Sub MergeCards(PS, PV, CS, CV, j, k)
        Dim i
        For i = j to k
            PS(i+2) = CS(i)
            PV(i+2) = CV(i)
        Next
    End Sub

    Sub CompareHands(Comb, PV, PS, k, BestScore, BestDesc)
        Dim i
        Dim j
        For i = 0 to k
            For j = 0 to 4
                iCheckVals(j) = PV(Mid(Comb(i), (j+1), 1))
                sCheckSuits(j) = PS(Mid(Comb(i), (j+1), 1))
            Next
            CheckHand iCheckVals, sCheckSuits, 4, BestScore, BestDesc
            ClearHandVars
        Next
    End Sub

    Sub CheckHand(CV, CS, j, BestScore, BestDesc)
        Dim i
        CheckHighCard CV, j 
        CheckPair CV, j
        If iRound > 0 Then
            CheckTwoPair CV, j
            CheckToaK CV, j
            CheckFoaK CV, j
            CheckStraight CV
            CheckFlush CS, CV
            CheckFullHouse CV
            CheckStraightFlush CV, CS
        End If
        If iCheckScore > BestScore Then
            BestScore = iCheckScore
            BestDesc = sCheckHandDesc
        End If
    End Sub

    Sub CheckHighcard(CV, j)
        Dim High
        High = ValToEng(CV(j))
        HasHC CV(j), High, CV
    End Sub

    Sub CheckPair(CV, j)
        Dim Eng
        Dim i
        For i = 0 to (j-1)
            If CV(i) = CV(i+1) And CV(i) > iPair(0) Then
                Eng = ValToEng(CV(i))
                HasPair CV(i), CV(i), Eng, CV
            End If
        Next
    End Sub

    Sub CheckTwoPair(CV, j)
        Dim HighEng
        Dim LowEng
        Dim i
        For i = 0 to (j-1)
            If CV(i) = CV(i+1) And CV(i) <> iPair(0) Then
                iPair(1) = CV(i)
                If iPair(0) > iPair(1) Then
                    HighEng = ValToEng(iPair(0))
                    LowEng = ValToEng(iPair(1))
                    HasTwoPair iPair(0), iPair(1), HighEng, LowEng, CV
                Else
                    HighEng = ValToEng(iPair(1))
                    LowEng = ValToEng(iPair(0))
                    HasTwoPair iPair(1), iPair(0), HighEng, LowEng, CV
                End If
            End If
        Next
    End Sub

    Sub CheckToaK(CV, j)
        Dim Eng
        Dim i
        For i = 0 to (j-2)
            If CV(i) = CV(i+1) And CV(i+1) = CV(i+2) Then
                Eng = ValToEng(CV(i))
                HasToaK CV(i), CV(i), Eng, CV
            End If
        Next
    End Sub

    Sub CheckFoaK(CV, j)
        Dim Eng
        Dim i
        For i = 0 to (j-3)
            If CV(i) = CV(i+1) And CV(i) = CV(i+2) And CV(i) = CV(i+3) Then
                Eng = ValToEng(CV(i))
                HasFoaK CV(i), CV(i), CV
            End If
        Next
    End Sub

    Sub CheckStraight(CV)
        Dim High
        Dim Low
        If (CV(0)+1) = CV(1) And (CV(1)+1) = CV(2) And (CV(2)+1) = CV(3) And (CV(3)+1) = CV(4) _
        Or (CV(0)+1) = CV(1) And (CV(1)+1) = CV(2) And (CV(2)+1) = CV(3) And (CV(3)+9) = CV(4) Then
            If CV(0) = 2 And CV(4) = 14 Then
                High = ValToEng(CV(3))
                Low = ValToEng(CV(4))
                HasStraight CV(3), High, Low, CV
            Else
                High = ValToEng(CV(4))
                Low = ValToEng(CV(0))
                HasStraight CV(4), High, Low, CV
            End If
        End If
    End Sub
    
    Sub CheckFlush(CS, CV)
        Dim High
        Dim Suit
        High = ValToEng(CV(4))
        If CS(0) = CS(1) And CS(0) = CS(2) And CS(0) = CS(3) And CS(0) = CS(4) Then
            Suit = SuitToEng(CS(0))
            HasFlush CV(4), High, Suit, CV
        End If
    End Sub

    Sub CheckFullHouse(CV)
        Dim Pair
        Dim Three
        If iPair(0) <> 0 And iPair(1) <> 0 And iThreeOfAKind <> 0 Then
            Three = ValToEng(iThreeOfAKind)
            If iPair(0) = iThreeOfAKind Then
                Pair = ValToEng(iPair(1))
                HasFH iThreeOfAKind, iPair(1), Three, Pair, CV
            Else
                Pair = ValToEng(iPair(0))
                HasFH iThreeOfAKind, iPair(0), Three, Pair, CV
            End If
        End If
    End Sub

    Sub CheckStraightFlush(CV, CS)
        Dim High
        Dim Suit
        If bStraight = True And bFlush = True Then
            Suit = SuitToEng(CS(0))
            High = ValToEng(CV(4))
            HasSF CV(4), High, Suit, CV
        End If
    End Sub

    Sub HasHC(Val, High, CV)
        sCheckHandDesc = sHandRes & " High Card " & High & "."
        sHighcard = High
        iCheckScore = CalcScore(1, Val) + CalcKickerScore(CV)
    End Sub

    Sub HasPair(Val, FstPair, Eng, CV)
        iPair(0) = FstPair
        sCheckHandDesc = sHandRes & " a Pair of " & Eng & "s."
        iCheckScore = CalcScore(10, Val) + CalcKickerScore(CV)
    End Sub

    Sub HasTwoPair(HighVal, LowVal, HighEng, LowEng, CV)
        sCheckHandDesc = sHandRes & " Two Pairs " & HighEng & "s and " & LowEng & "s."
        iCheckScore = (CalcScore(100, HighVal) + CalcScore(1, LowVal)) + CalcKickerScore(CV)
    End Sub
    
    Sub HasToaK(Val, ToaK, Eng, CV)
        sCheckHandDesc = sHandRes & " Three of a Kind - " & Eng & "s."
        iCheckScore = CalcScore(1000, Val) + CalcKickerScore(CV)
        iThreeOfAKind = ToaK
    End Sub

    Sub HasFoaK(Val, Eng, CV)
        sCheckHandDesc = sHandRes & " Four of a Kind - " & Eng & "s."
        iCheckScore = CalcScore(10000000, Val) + CalcKickerScore(CV)
    End Sub

    Sub HasStraight(Val, HighEng, LowEng, CV)
        sCheckHandDesc = sHandRes & " a Straight from " & LowEng & " to " & HighEng & "."
        iCheckScore = CalcScore(10000, Val) + CalcKickerScore(CV)
        bStraight = True
    End Sub

    Sub HasFlush(Val, High, Suit, CV)
        sCheckHandDesc = sHandRes & " a " & High & " high Flush of " & Suit & "."
        iCheckScore = CalcScore(100000, Val) + CalcKickerScore(CV)
        bFlush = True
    End Sub

    Sub HasFH(TVal, PVal, Three, Pair, CV)
        sCheckHandDesc = sHandRes & " a Full House - " & Three & "s full of " & Pair & "s."
        iCheckScore = (CalcScore(1000000, TVal) +  CalcScore(1, PVal)) + CalcKickerScore(CV)
    End Sub

    Sub HasSF(Val, High, Suit, CV)
        sCheckHandDesc = sHandRes & " a " & High & " high Straight-Flush of " & Suit & "."
        iCheckScore = CalcScore(100000000, Val) + CalcKickerScore(CV)
    End Sub

    Function CalcScore(Multi, Val)
        CalcScore = (Multi * Val)
    End Function

    Function CalcKickerScore(CV)
        CalcKickerScore = ((0.01 * CV(4)) + (0.0001 * CV(3)) + (0.000001 * CV(2)) + (0.00000001 * CV(1)) + (0.0000000001 * CV(0)))
    End Function

    Function ValToEng(i)
        If i = 14 Then
            ValToEng = "Ace"
        ElseIf i = 13 Then
            ValToEng = "King"
        ElseIf i = 12 Then
            ValToEng = "Queen"
        ElseIf i = 11 Then
            ValToEng = "Jack"
        ElseIf i = 10 Then
            ValToEng = "Ten"
        ElseIf i = 9 Then
            ValToEng = "Nine"
        ElseIf i = 8 Then
            ValToEng = "Eight"
        ElseIf i = 7 Then
            ValToEng = "Seven"
        ElseIf i = 6 Then
            ValToEng = "Six"
        ElseIf i = 5 Then
            ValToEng = "Five"
        ElseIf i = 4 Then
            ValToEng = "Four"
        ElseIf i = 3 Then
            ValToEng = "Three"
        ElseIf i = 2 Then
            ValToEng = "Two"
        End If
    End Function

    Function SuitToEng(a)
        If a = "C" Then
            SuitToEng = "Clubs"
        ElseIf a = "D" Then
            SuitToEng = "Diamonds"
        ElseIf a = "H" Then
            SuitToEng = "Hearts"
        ElseIf a = "S" Then
            SuitToEng = "Spades"
        End If
    End Function

    Sub Winner()
        Dim SplitPot
        If iPBestScore > iC1BestScore Then
            sHandRes = "You win the pot with"
            sPBestHandDesc = sHandRes & sPBestHandDesc
            ShowBest sPBestHandDesc
            DistributePot iPChips, iPot
        ElseIf iPBestScore < iC1BestScore Then
            sHandRes = "Computer 1 wins the pot with"
            sC1BestHandDesc = sHandRes & sC1BestHandDesc
            ShowBest sC1BestHandDesc
            DistributePot iC1Chips, iPot
        Else
            sHandRes = "You split the pot with"
            sPBestHandDesc = sHandRes & sPBestHandDesc
            ShowBest sPBestHandDesc
            SplitPot = Int(iPot/2)
            iPot = Int(iPot/2)
            DistributePot iPChips, SplitPot
            DistributePot iC1Chips, iPot
        End If
    End Sub
    
    Sub ShowBest(Desc)
        parResult.innerHTML = parResult.innerHTML & Desc
    End Sub