    Sub StartEndLock(Change1, Change2)
        Dim i
        btnFold.disabled = Change1
        For i = 0 to 5
            document.getElementById("btnBet" & i).disabled = Change1
        Next
        btnStart.disabled = Change2
    End Sub

    Sub Blinds()
        CalcBlind
        Human.iRoundBet = (iAnte + iBlind)
        CheckLegal Human.iRoundBet, Human.iChipCount, Comp1.iChipCount
        Comp1.iRoundBet = Human.iRoundBet
        Comp1.iChipCount = (Comp1.iChipCount - Comp1.iRoundBet)
        MoveToPot
        ShowBet
    End Sub

    Sub CalcBlind()
        iBlind = ((Int(iHandNum/10)+1)*50)
        iAnte = ((Int(iHandNum/20)-1)*25)
        If iAnte < 0 Then
            iAnte = 0
        End If
        If (iHandNum/10) = Int(iHandNum/10) Then
            Alert("Blinds are up: " & iBlind & "big blind and " & iAnte & " ante.")
        End If
    End Sub
    
    Sub PlayerBet(GreatScore, VeryHighScore, HighScore, MedScore, LowScore)
        If iPChoice < 5 Then
            Human.iRoundBet = (iPChoice * iBlind)
            CheckLegal Human.iRoundBet, Human.iChipCount, Comp1.iChipCount
            CompMove GreatScore, VeryHighScore, HighScore, MedScore, LowScore, 0
        ElseIf iPChoice = 5 Then
            Human.iRoundBet = Human.iChipCount
            CheckLegal Human.iRoundBet, Human.iChipCount, Comp1.iChipCount
            CompMove GreatScore, VeryHighScore, HighScore, MedScore, LowScore, 0
        End If
    End Sub

    Sub CheckLegal(Bet, OwnChips, OppChips)
        If Bet > OwnChips Then
            Bet = OwnChips
        ElseIf Bet > OppChips Then
            Bet = OppChips
        End If
        OwnChips = (OwnChips - Bet)
    End Sub

    Sub CompMove(GreatScore, VeryHighScore, HighScore, MedScore, LowScore, PayBlind)
        bC1Moved = True
        If Comp1.iBestScore => GreatScore And iPChoice < 5 Then
            CompRaise 5
        ElseIf Comp1.iBestScore => VeryHighScore And iPChoice < 4 Then
            CompRaise 4
        ElseIf Comp1.iBestScore => HighScore And iPChoice < 3 Then
            CompRaise 3
        ElseIf Comp1.iBestScore => MedScore And iPChoice < 2 Then
            CompRaise 2
        ElseIf Comp1.iBestScore => LowScore And iPChoice < 1 Then
            CompRaise 1
        ElseIf Comp1.iBestScore => GreatScore And iPChoice = 5 Or Comp1.iBestScore => VeryHighScore And iPChoice = 4 Or Comp1.iBestScore => HighScore And iPChoice = 3 _
        Or Comp1.iBestScore => MedScore And iPChoice = 2 Or Comp1.iBestScore => LowScore And iPChoice = 1 Or Comp1.iBestScore => 0 And iPChoice = 0 Then
            CompCall
            ShowBet
        Else
            CompFold
        End If
    End Sub

    Sub CompRaise(CompChoice)
        RaiseLock "disabled"
        If CompChoice = 5 Then
            Comp1.iRoundBet = Comp1.iChipCount
            CheckLegal Comp1.iRoundBet, Comp1.iChipCount, Human.iChipCount
            Alert("Computer 1 goes all in; " & (Comp1.iRoundBet - Human.iRoundBet) & " chips to call.")
        Else
            Comp1.iRoundBet = (Human.iRoundBet + (CompChoice * iBlind))
            CheckLegal Comp1.iRoundBet, Comp1.iChipCount, Human.iChipCount
            Alert("Computer 1 raises to " & Comp1.iRoundBet & "; " & (Comp1.iRoundBet - Human.iRoundBet) & " chips to call.")
        End If
    End Sub

    Sub RaiseLock(Change)
        Dim i
        For i = 1 to 5
            document.getElementById("btnBet" & i).disabled = change
        Next
    End Sub

    Sub PlayerCall()
        RaiseLock ""
        bC1Moved = False
        Human.iChipCount = (Human.iChipCount - (Comp1.iRoundBet - Human.iRoundBet))
        Human.iRoundBet = Comp1.iRoundBet
        MoveToPot
        ShowBet
        iRound = (iRound+1)
    End Sub

    Sub CompCall()
        If Human.iRoundBet = 0 Then
            Alert("Computer 1 checks.")
        Else
            Alert("Computer 1 calls your bet of " & Human.iRoundBet & " chips.")
        End If
        Comp1.iRoundBet = Human.iRoundBet
        Comp1.iChipCount = (Comp1.iChipCount - Comp1.iRoundBet)
        MoveToPot
        bC1Moved = False
        If iRound = 0 Then
            iRound = (iRound+1)
            Flop
        ElseIf iRound = 1 Then
            iRound = (iRound+1)
            Turn
        ElseIf iRound = 2 Then
            iRound = (iRound+1)
            River
        ElseIf iRound = 3 Then
            HandOver
        End If
    End Sub

    Sub CompFold()
        MoveToPot
        Fold Human.iChipCount, "Computer 1 folds the pot; You win."
        bC1Fold = True
    End Sub

    Sub Fold(WinnerChips, Desc)
        StartEndLock "Disabled", ""
        ShowBet
        ShowBest Desc
        DistributePot WinnerChips, iPot
    End Sub

    Sub MoveToPot()
        iPot = (iPot + Human.iRoundBet + Comp1.iRoundBet)
        Human.iRoundBet = 0
        Comp1.iRoundBet = 0
        iPChoice = 0
    End Sub

    Sub ShowBet()
        parPChips.innerHTML = "<Table><Tr><Th>Player</Th></Tr></Table><Table><Tr><Td>Chips:</Td><Td>" & Human.iChipCount & "</Td></Tr></Table>" & Human.iBestScore
        parC1Chips.innerHTML = "<Table><Tr><Th>Computer 1</Th></Tr></Table><Table><Tr><Td>Chips:</Td><Td>" & Comp1.iChipCount & "</Td></Tr></Table>" & Comp1.iBestScore
        parResult.innerHTML = "<Table><Tr><Td>Current Pot:</Td><Td>" & iPot & "</Td></Tr></Table>"
    End Sub

    Sub DistributePot(WinnerChips, Chips)
        WinnerChips = (WinnerChips + Chips)
        Chips = 0
        iRound = 0
    End Sub