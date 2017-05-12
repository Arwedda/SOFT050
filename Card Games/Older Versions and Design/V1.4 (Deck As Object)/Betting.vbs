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
        iPRoundBet = (iAnte + iBlind)
        CheckLegal iPRoundBet, iPChips, iC1Chips
        iC1RoundBet = iPRoundBet
        iC1Chips = (iC1Chips - iC1RoundBet)
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
            iPRoundBet = (iPChoice * iBlind)
            CheckLegal iPRoundBet, iPChips, iC1Chips
            CompMove GreatScore, VeryHighScore, HighScore, MedScore, LowScore, 0
        ElseIf iPChoice = 5 Then
            iPRoundBet = iPChips
            CheckLegal iPRoundBet, iPChips, iC1Chips
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
        If iC1BestScore => GreatScore And iPChoice < 5 Then
            CompRaise 5
        ElseIf iC1BestScore => VeryHighScore And iPChoice < 4 Then
            CompRaise 4
        ElseIf iC1BestScore => HighScore And iPChoice < 3 Then
            CompRaise 3
        ElseIf iC1BestScore => MedScore And iPChoice < 2 Then
            CompRaise 2
        ElseIf iC1BestScore => LowScore And iPChoice < 1 Then
            CompRaise 1
        ElseIf iC1BestScore => GreatScore And iPChoice = 5 Or iC1BestScore => VeryHighScore And iPChoice = 4 Or iC1BestScore => HighScore And iPChoice = 3 _
        Or iC1BestScore => MedScore And iPChoice = 2 Or iC1BestScore => LowScore And iPChoice = 1 Or iC1BestScore => 0 And iPChoice = 0 Then
            CompCall
            ShowBet
        Else
            CompFold
        End If
    End Sub

    Sub CompRaise(CompChoice)
        RaiseLock "disabled"
        If CompChoice = 5 Then
            iC1RoundBet = iC1Chips
            CheckLegal iC1RoundBet, iC1Chips, iPChips
            Alert("Computer 1 goes all in; " & (iC1RoundBet - iPRoundBet) & " chips to call.")
        Else
            iC1RoundBet = (iPRoundBet + (CompChoice * iBlind))
            CheckLegal iC1RoundBet, iC1Chips, iPChips
            Alert("Computer 1 raises to " & iC1Roundbet & "; " & (iC1RoundBet - iPRoundBet) & " chips to call.")
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
        iPChips = (iPChips - (iC1RoundBet - iPRoundBet))
        iPRoundBet = iC1RoundBet
        MoveToPot
        ShowBet
        iRound = (iRound+1)
    End Sub

    Sub CompCall()
        If iPRoundBet = 0 Then
            Alert("Computer 1 checks.")
        Else
            Alert("Computer 1 calls your bet of " & iPRoundBet & " chips.")
        End If
        iC1RoundBet = iPRoundBet
        iC1Chips = (iC1Chips - iC1RoundBet)
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
        Fold iPChips, "Computer 1 folds the pot; You win."
        bC1Fold = True
    End Sub

    Sub Fold(WinnerChips, Desc)
        StartEndLock "Disabled", ""
        ShowBet
        ShowBest Desc
        DistributePot WinnerChips, iPot
    End Sub

    Sub MoveToPot()
        iPot = (iPot + iPRoundBet + iC1RoundBet)
        iPRoundBet = 0
        iC1Roundbet = 0
        iPChoice = 0
    End Sub

    Sub ShowBet()
        parPChips.innerHTML = "<Table><Tr><Th>Player</Th></Tr></Table><Table><Tr><Td>Chips:</Td><Td>" & iPChips & "</Td></Tr></Table>" & iPBestScore
        parC1Chips.innerHTML = "<Table><Tr><Th>Computer 1</Th></Tr></Table><Table><Tr><Td>Chips:</Td><Td>" & iC1Chips & "</Td></Tr></Table>" & iC1BestScore
        parResult.innerHTML = "<Table><Tr><Td>Current Pot:</Td><Td>" & iPot & "</Td></Tr></Table>"
    End Sub

    Sub DistributePot(WinnerChips, Chips)
        WinnerChips = (WinnerChips + Chips)
        Chips = 0
        iRound = 0
    End Sub