    Sub FreshHand()
        Hand(0) = ""
        Hand(1) = ""
        Hand(2) = ""
        Hand(3) = ""
        Hand(4) = ""
        HighCard = ""
        Pair(0) = ""
        Pair(1) = ""
        ThreeOfAKind = ""
        Straight = ""
        Flush = ""
        parHand.innerText = ""
    End Sub

    Sub PickCards()
        Do While Hand(0) = Hand(1) or Hand(0) = Hand(2) or Hand(0) = Hand(3) or Hand(0) = Hand(4) or Hand(1) = Hand(2) or Hand(1) = Hand(3) or Hand(1) = Hand(4) or Hand(2) = Hand(3) or Hand(2) = Hand(4) or Hand(3) = Hand(4)
            Hand(0) = Int(RND() * 52)
            Hand(1) = Int(RND() * 52)
            Hand(2) = Int(RND() * 52)
            Hand(3) = Int(RND() * 52)
            Hand(4) = Int(RND() * 52)
        Loop
    End Sub

    Sub Sort()
        Dim Swap
        Dim Count
        Dim Temp
        Do
            Swap = False
            For Count = 0 to 3
                If CardValue(Count) > CardValue(Count + 1) Then
                    Temp = CardValue(Count + 1)
                    CardValue(Count + 1) = CardValue(Count)
                    CardValue(Count) = Temp
                    Temp = Deck(Hand(Count + 1))
                    Deck(Hand(Count + 1)) = Deck(Hand(Count))
                    Deck(Hand(Count)) = Temp
                    Swap = True
                End If
            Next
        Loop Until Swap = False
    End Sub

    Sub SuitsAndValues()
    Dim i
        For i = 0 to 4
            CardSuit(i) = Left(Deck(Hand(i)), 1)
            CardValue(i) = CINT(Mid(Deck(Hand(i)), 2, 2))
            document.getElementById("imgCard" & i).src = Deck(Hand(i))
        Next
    End Sub

    'Calculating Hands

    Sub CheckStraightFlush()
        If Straight = "Yes" And Flush = "Yes" Then
            parHand.innerText = "You have a Straight Flush of " & CardSuit(0) & "s from " & CardValue(0) & " to " & CardValue(4) & "."
        ElseIf Straight = "Royal" and Flush = "Yes" Then
             parHand.innerText = "You have a Royal Straight Flush - " & CardSuit(0) & "s from " & CardValue(1) & " to " & CardValue(0) & "."           
        End If
    End Sub

    Sub CheckFourOfAKind()
        If CardValue(1) = CardValue(2) And CardValue(1) = CardValue(3) And CardValue(1) = CardValue(4) Then
            parHand.innerText = "You have Four of a Kind - " & CardValue(1) & "s."
        ElseIf CardValue(0) = CardValue(1) And CardValue(0) = CardValue(2) And CardValue(0) = CardValue(3) Then
            parHand.innerText = "You have Four of a Kind - " & CardValue(0) & "s."
        End If
    End Sub

    Sub CheckFullHouse()
        If Pair(0) <> "" And Pair(1) <> "" And ThreeOfAKind <> "" Then
            If Pair(0) = ThreeOfAKind Then
                parHand.innerText = "You have a Full House - " & ThreeOfAKind & "s full of " & Pair(1) & "s."
            Else
                parHand.innerText = "You have a Full House - " & ThreeOfAKind & "s full of " & Pair(0) & "s."
            End If
        End If
    End Sub
    
    Sub CheckFlush()
        If CardSuit(0) = CardSuit(1) And CardSuit(0) = CardSuit(2) And CardSuit(0) = CardSuit(3) And CardSuit(0) = CardSuit(4) Then
            parHand.innerText = "You have a " & HighCard & " high Flush of " & CardSuit(0) & "s."
            Flush = "Yes"
        End If
    End Sub

    Sub CheckStraight()
        If CardValue(0) + 9 = CardValue(1) And CardValue(0) + 10 = CardValue(2) And CardValue(0) + 11 = CardValue(3) And CardValue(0) + 12 = CardValue(4) Then
            parHand.innerText = "You have a Straight from " & CardValue(1) & " to " & CardValue(0) & "."
            Straight = "Royal"
        ElseIf CardValue(0) + 1 = CardValue(1) And CardValue(0) + 2 = CardValue(2) And CardValue(0) + 3 = CardValue(3) And CardValue(0) + 4 = CardValue(4) Then
            parHand.innerText = "You have a Straight from " & CardValue(0) & " to " & CardValue(4) & "."
            Straight = "Yes"
        End If
    End Sub

    Sub CheckThreeOfAKind()
        If CardValue(0) = CardValue(1) And CardValue(0) = CardValue(2) Then
            parHand.innerText = "You have Three Of A Kind - " & CardValue(0) & "s."
            ThreeOfAKind = CardValue(0)
        ElseIf CardValue(1) = CardValue(2) And CardValue(1) = CardValue(3) Then
            parHand.innerText = "You have Three Of A Kind - " & CardValue(1) & "s."
            ThreeOfAKind = CardValue(1)
        ElseIf CardValue(2) = CardValue(3) And CardValue(2) = CardValue(4) Then
            parHand.innerText = "You have Three Of A Kind - " & CardValue(2) & "s."
            ThreeOfAKind = CardValue(2)
        End If   
    End Sub

    Sub CheckTwoPair(a, b, c)
        If CardValue(a) = CardValue(b) Then
            Pair(1) = CardValue(a)
            parHand.innerText = "You have Two Pairs " & Pair(1) & "s and " & Pair(0) & "s."
        ElseIf CardValue(b) = CardValue(c) Then
            Pair(1) = CardValue(b)
            parHand.innerText = "You have Two Pairs " & Pair(1) & "s and " & Pair(0) & "s."
        End If
    End Sub

    Sub CheckPair()
        If CardValue(0) = CardValue(1) Then
            Pair(0) = CardValue(0)
            parHand.innerText = "You have a Pair of " & Pair(0) & "s."
            CheckThreeOfAKind
            CheckTwoPair 2, 3, 4
        ElseIf CardValue(1) = CardValue(2) Then
            Pair(0) = CardValue(1)
            parHand.innerText = "You have a Pair of " & Pair(0) & "s."
            CheckThreeOfAKind
            CheckTwoPair 0, 3, 4
        ElseIf CardValue(2) = CardValue(3) Then
            Pair(0) = CardValue(2)
            parHand.innerText = "You have a Pair of " & Pair(0) & "s."
            CheckThreeOfAKind
            CheckTwoPair 0, 1, 4
        ElseIf CardValue(3) = CardValue(4) Then
            Pair(0) = CardValue(3)
            parHand.innerText = "You have a Pair of " & Pair(0) & "s."
            CheckThreeOfAKind
            CheckTwoPair 0, 1, 2
        End If
    End Sub

    Sub CheckHighCard()
        If CardValue(0) = " 1" Then
            HighCard = "Ace"
            parHand.innerText = "You have High Card " & HighCard & "."
        ElseIf CardValue(4) = "13" Then
            HighCard = "King"
            parHand.innerText = "You have High Card " & HighCard & "."
        ElseIf CardValue(4) = "12" Then
            HighCard = "Queen"
            parHand.innerText = "You have High Card " & HighCard & "."
        ElseIf CardValue(4) = "11" Then
            HighCard = "Jack"
            parHand.innerText = "You have High Card " & HighCard & "."
        Else
            HighCard = CardValue(4)
            parHand.innerText = "You have High Card " & HighCard & "."
        End If
    End Sub