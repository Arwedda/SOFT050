    Sub FreshDeck()
        sDeck = Array("C14.png", "C 2.png", "C 3.png", "C 4.png", "C 5.png", "C 6.png", "C 7.png","C 8.png", "C 9.png", "C10.png", "C11.png", "C12.png", "C13.png", _
        "D14.png", "D 2.png", "D 3.png", "D 4.png", "D 5.png", "D 6.png", "D 7.png", "D 8.png", "D 9.png", "D10.png", "D11.png", "D12.png", "D13.png", _
        "H14.png", "H 2.png", "H 3.png", "H 4.png", "H 5.png", "H 6.png", "H 7.png", "H 8.png", "H 9.png", "H10.png", "H11.png", "H12.png", "H13.png", _
        "S14.png", "S 2.png", "S 3.png", "S 4.png", "S 5.png", "S 6.png", "S 7.png", "S 8.png", "S 9.png", "S10.png", "S11.png", "S12.png", "S13.png")
    End Sub

    Sub CreateCombinations()
        sSixCardCombos = Array("01234", "01235", "01245", "01345", "02345", 12345)  'Card combinations to check at the Turn (penultimate card)
        sSevCardCombos = Array("01234", "01235", "01245", "01345", "02345", 12345, "01236","01246", "01256", "01346", "01356", "01456", "02346", _
        "02356", "02456", "03456", 12346, 12356, 12456, 13456, 23456)   'Card combinations to check at the River (final card)
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
        HMid = (document.body.clientHeight - 250) / 2   'Centre height of card = centre of screen
        WMid = (document.body.clientWidth - 200) / 2    'Centre width of card = centre of screen
        For i = 0 to 4  'Positioning community cards
            document.getElementById("imgCC" & i).style.posTop = HMid
            document.getElementById("imgCC" & i).style.posLeft = WMid + (210 * (i-2)) 
        Next
        parResult.style.posTop = HMid   'Positioning feedback paragraph

        HMid = (document.body.clientHeight - 250)   'Bottom of card = bottom of screen
        For i = 0 to 1  'Positioning player hole cards
            document.getElementById("imgPC" & i).style.posTop = HMid
        Next
        imgPC0.style.posLeft = WMid - 105
        imgPC1.style.posLeft = WMid + 105
        parPChips.style.posTop = HMid
        parPChips.style.posLeft = WMid + 310    'To the right of cards

        For i = 0 to 1  'Position computer 1 hole cards
            document.getElementById("imgC1C" & i).style.posTop = 0  'Top of card = top of screen
        Next
        imgC1C0.style.posLeft = WMid - 105
        imgC1C1.style.posLeft = WMid + 105
        parC1Chips.style.posTop = 0
        parC1Chips.style.posLeft = WMid + 310   'To the right of cards

        ShowBet
    End Sub
        
    Sub FreshHand()
        ClearCommunalVars
        ClearPlayerVars
        ClearCompVars
    End Sub

    Sub ClearCommunalVars()
        Dim i
        iHandNum = iHandNum + 1
        iRound = 0
        iPot = 0
        ClearHandVars
        ClearCards sCCard, iCValue, sCSuit, "imgCC", 4
        ClearScores iCheckScore 
        sHandRes = "Your current best hand is"
        sCheckHandDesc = ""
        For i = 0 to 4
            iCheckVals(i) = 0
            sCheckSuits(i) = ""
        Next
    End Sub
    
    Sub ClearPlayerVars()
        ClearCards sPCard, iPValue, sPSuit, "imgPC", 1
        ClearScores iPBestScore
        sPBestHandDesc = ""
    End Sub

    Sub ClearCompVars()
        ClearCards sC1Card, iC1Value, sC1Suit, "imgC1C", 1
        ClearScores iC1BestScore
        sC1BestHandDesc = ""
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
        Scr = 0
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
        Do While sPCard(0) = sPCard(1)  'Player
            PickCards sPCard, 0, 1
        Loop
	    RemoveFromDeck sPCard, 0, 1
        Do While sC1Card(0) = sC1Card(1) Or sC1Card(0) = "" Or sC1Card(1) = ""  'Computer 1
            PickCards sC1Card, 0, 1
        Loop
	    RemoveFromDeck sC1Card, 0, 1
    End Sub

    Sub DealFlop()
        Do While sCCard(0) = sCCard(1) Or sCCard(0) = sCCard(2)  Or sCCard(1) = sCCard(2) Or sCCard(0) = "" Or sCCard(1) = "" Or sCCard(2) = ""
            PickCards sCCard, 0, 2
        Loop
	    RemoveFromDeck sCCard, 0, 2
    End Sub

    Sub DealTurn()
        Do While sCCard(3) = ""
           PickCards sCCard, 3, 3
        Loop
	    RemoveFromDeck sCCard, 3, 3
    End Sub

    Sub DealRiver()
        Do While sCCard(4) = ""
            PickCards sCCard, 4, 4
        Loop
    End Sub

    Sub PickCards(Card, j, k)
        Dim i
        For i = j to k
            Card(i) = sDeck(Int(RND()*52))  'Pull random image from array and deal to a player or communal cards
        Next
    End Sub

    Sub RemoveFromDeck(Card, k, l)
        Dim Found
        Dim i
        Dim j
        For i = k to l  'Cycles through each dealt cards
            j = 0
            Do  'Cycles through deck until it finds the dealt card
                Found = False
                If Card(i) = sDeck(j) Then  'When cards match then remove from deck so it can't be dealt again this hand
                    Found = True
                    sDeck(j) = ""
                End If
                j = (j+1)
            Loop Until Found = True
        Next
    End Sub

    Sub SuitsAndValues(CV, CS, pic, Cards, j, k)
        Dim i
        For i = j to k
            CS(i) = Left(Cards(i), 1)   'Extract single letter from image name to find suit
            CV(i) = CINT(Mid(Cards(i), 2, 2))   'Extract 1-2 numbers from image name to find value
        Next
    End Sub

    Sub Sort(CV, CS, Cards, j, ShowPics)
        Dim Swap
        Dim Count
        Dim Temp
        Do  'Continue until previous round didn't require cards to swap places
            Swap = False
            For Count = 0 to j  'Loop through cards to be sorted
                If CV(Count) > CV(Count+1) Then 'If current card has higher value than the next card then swap the suits and values
                    Swap = True
                    Temp = CV(Count+1)
                    CV(Count+1) = CV(Count)
                    CV(Count) = Temp
                    Temp = CS(Count+1)
                    CS(Count+1) = CS(Count)
                    CS(Count) = Temp
                    If ShowPics = True Then 'Only swap the images to be value order when it makes sense to do so (i.e. for the flop)
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
        For i = j to k  'Cycle through dealt communal cards' values and suits and merge them with individual player's hole cards
            PS(i+2) = CS(i)
            PV(i+2) = CV(i)
        Next
    End Sub

    Sub CompareHands(Comb, PV, PS, k, BestScore, BestDesc)
        Dim i
        Dim j
        For i = 0 to k  'For however many hand combinations are currently available
            For j = 0 to 4  'Extract the relevant suits and values using combinations array
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
        If iRound > 0 Then  'Only need to check for these hands if flop has been dealt
            CheckTwoPair CV, j
            CheckToaK CV, j
            CheckFoaK CV, j
            CheckStraight CV
            CheckFlush CS, CV
            CheckFullHouse CV
            CheckStraightFlush CV, CS
        End If
        If iCheckScore > BestScore Then 'Update best score
            BestScore = iCheckScore
            BestDesc = sCheckHandDesc
        End If
    End Sub

    Sub CheckHighcard(CV, j)
        Dim High
        High = ValToEng(CV(j))  'Right-most card will be highest
        HasHC CV(j), High, CV
    End Sub

    Sub CheckPair(CV, j)
        Dim Eng
        Dim i
        For i = 0 to (j-1)
            If CV(i) = CV(i+1) And CV(i) > iPair(0) Then    'If two adjacent cards have the same value you have a pair
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
            If CV(i) = CV(i+1) And CV(i) <> iPair(0) Then   'If there is a pair that doesn't match previous pair
                iPair(1) = CV(i)
                If iPair(0) > iPair(1) Then 'Highest pair has the highest multiplier for hand score
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
            If CV(i) = CV(i+1) And CV(i+1) = CV(i+2) Then   'If three adjacent cards have the same value you have three of a kind
                Eng = ValToEng(CV(i))
                HasToaK CV(i), CV(i), Eng, CV
            End If
        Next
    End Sub

    Sub CheckFoaK(CV, j)
        Dim Eng
        Dim i
        For i = 0 to (j-3)
            If CV(i) = CV(i+1) And CV(i) = CV(i+2) And CV(i) = CV(i+3) Then 'If four adjacent cards have the same value you have four of a kind
                Eng = ValToEng(CV(i))
                HasFoaK CV(i), Eng, CV
            End If
        Next
    End Sub

    Sub CheckStraight(CV)
        Dim High
        Dim Low
        If (CV(0)+1) = CV(1) And (CV(1)+1) = CV(2) And (CV(2)+1) = CV(3) And (CV(3)+1) = CV(4) _
        Or (CV(0)+1) = CV(1) And (CV(1)+1) = CV(2) And (CV(2)+1) = CV(3) And (CV(3)+9) = CV(4) Then 'If card values are x to x+4 or 14 and 2 to 4 you have a straight
            If CV(0) = 2 And CV(4) = 14 Then    'If straight is ace to 5
                High = ValToEng(CV(3))
                Low = ValToEng(CV(4))
                HasStraight CV(3), High, Low, CV
            Else    'Otherwise it is a "normal" straight
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
        If CS(0) = CS(1) And CS(0) = CS(2) And CS(0) = CS(3) And CS(0) = CS(4) Then 'If all 5 card suits are the same you have a flush
            Suit = SuitToEng(CS(0))
            HasFlush CV(4), High, Suit, CV
        End If
    End Sub

    Sub CheckFullHouse(CV)
        Dim Pair
        Dim Three
        If iPair(0) <> 0 And iPair(1) <> 0 And iThreeOfAKind <> 0 Then  'If you have two pairs and one of them is actually three of a kind then you a full house
            Three = ValToEng(iThreeOfAKind)
            If iPair(0) = iThreeOfAKind Then    'Works out which pair is actually three of a kind for higher score multiplier
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
        If bStraight = True And bFlush = True Then  'If we have found a straight and a flush within the same 5 card combination then we have a straight flush
            Suit = SuitToEng(CS(0))
            High = ValToEng(CV(4))
            HasSF CV(4), High, Suit, CV
        End If
    End Sub

    Sub HasHC(Val, High, CV)    'Subroutines to construct a hand score to calculate which hand is the best and feedback to inform the player what the winning hand consists of
        sCheckHandDesc = sHandRes & " High Card " & High & "."
        sHighcard = High
        iCheckScore = CalcScore(1, Val) + CalcKickerScore(CV)
    End Sub

    Sub HasPair(Val, FstPair, Eng, CV)
        iPair(0) = FstPair
        If Eng = "Six" Then
            Eng = "Sixe"
        End If
        sCheckHandDesc = sHandRes & " a Pair of " & Eng & "s."
        iCheckScore = CalcScore(10, Val) + CalcKickerScore(CV)
    End Sub

    Sub HasTwoPair(HighVal, LowVal, HighEng, LowEng, CV)
        If HighEng = "Six" Then
            HighEng = "Sixe"
        ElseIf LowEng = "Six" Then
            LowEng = "Sixe"
        End If
        sCheckHandDesc = sHandRes & " Two Pairs " & HighEng & "s and " & LowEng & "s."
        iCheckScore = (CalcScore(100, HighVal) + CalcScore(1, LowVal)) + CalcKickerScore(CV)
    End Sub
    
    Sub HasToaK(Val, ToaK, Eng, CV)
        If Eng = "Six" Then
            Eng = "Sixe"
        End If
        sCheckHandDesc = sHandRes & " Three of a Kind - " & Eng & "s."
        iCheckScore = CalcScore(1000, Val) + CalcKickerScore(CV)
        iThreeOfAKind = ToaK
    End Sub

    Sub HasFoaK(Val, Eng, CV)
        If Eng = "Six" Then
            Eng = "Sixe"
        End If
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
        If HighEng = "Six" Then
            HighEng = "Sixe"
        ElseIf LowEng = "Six" Then
            LowEng = "Sixe"
        End If
        sCheckHandDesc = sHandRes & " a Full House - " & Three & "s full of " & Pair & "s."
        iCheckScore = (CalcScore(1000000, TVal) +  CalcScore(1, PVal)) + CalcKickerScore(CV)
    End Sub

    Sub HasSF(Val, High, Suit, CV)
        sCheckHandDesc = sHandRes & " a " & High & " high Straight-Flush of " & Suit & "."
        iCheckScore = CalcScore(100000000, Val) + CalcKickerScore(CV)
    End Sub

    Function CalcScore(Multi, Val)  'Score consists of a multiplier that grows with hand strength and the value of the highest card constructing said hand
        CalcScore = (Multi * Val)
    End Function

    Function CalcKickerScore(CV)    'Very small values that only affect who the winner is if they share the same "main" score
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
        If iPBestScore > iC1BestScore Then  'You have highest score
            sHandRes = "You win the pot with"
            sPBestHandDesc = sHandRes & Mid(sPBestHandDesc, 26)
            ShowBest sPBestHandDesc
            DistributePot iPChips, iPot
        ElseIf iPBestScore < iC1BestScore Then  'Computer 1 has highest score
            sHandRes = "Computer 1 wins the pot with"
            sC1BestHandDesc = sHandRes & Mid(sC1BestHandDesc, 26)
            ShowBest sC1BestHandDesc
            DistributePot iC1Chips, iPot
        Else    'You have the same score
            sHandRes = "You split the pot with"
            sPBestHandDesc = sHandRes & Mid(sPBestHandDesc, 26)
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