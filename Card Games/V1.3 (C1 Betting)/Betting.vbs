    Sub StartEndLock(Change1, Change2)  'Enables/Disables based on position in hand
        Dim i
        btnFold.disabled = Change1
        For i = 0 to 5
            document.getElementById("btnBet" & i).disabled = Change1
        Next
        btnStart.disabled = Change2
    End Sub

    Sub RaiseLock(Change)   'Stops player from re-raising to prevent raise loop between player and computer
        Dim i
        For i = 1 to 5
            document.getElementById("btnBet" & i).disabled = change
        Next
    End Sub

    Sub Blinds()
        CalcBlind
        iPRoundBet = (iAnte + iBlind)
        CheckLegal iPRoundBet, iPChips, iC1Chips
        iC1RoundBet = iPRoundBet
        iC1Chips = (iC1Chips - iC1RoundBet)
        MoveToPot
    End Sub

    Sub CalcBlind()
        iBlind = ((Int(iHandNum/10)+1)*50)  'Every 10 hands the blind goes up by 50
        iAnte = ((Int(iHandNum/20)-1)*20)   'Every 20 hands starting at hand 40 ante goes up by 20
        If iAnte < 0 Then   'Never negative ante
            iAnte = 0
        End If
        If (iHandNum/10) = Int(iHandNum/10) Then
            Alert("Blinds are up: " & iBlind & " chip Blind and " & iAnte & " chip ante.")
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
        If Bet > OwnChips Then  'Checks bet is below player and computer chips
            Bet = OwnChips
        ElseIf Bet > OppChips Then
            Bet = OppChips
        End If
        OwnChips = (OwnChips - Bet)
    End Sub

    Sub CompMove(GreatScore, VeryHighScore, HighScore, MedScore, LowScore, PayBlind)
        bC1Moved = True
        If iC1BestScore => GreatScore And iPChoice < 5 Then 'If player bet too low for computer hand then raise accordingly
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
        Or iC1BestScore => MedScore And iPChoice = 2 Or iC1BestScore => LowScore And iPChoice = 1 Or iC1BestScore => 0 And iPChoice = 0 Then    'If player bet right for computer hand then call
            CompCall
        Else    'Player bet too high for computer hand then fold
            CompFold
        End If

    End Sub

    Sub CompRaise(CompChoice)
        RaiseLock "disabled"
        If CompChoice = 5 Then  'Computer 1 all in
            iC1RoundBet = iC1Chips
            CheckLegal iC1RoundBet, iC1Chips, iPChips
            Alert("Computer 1 goes all in; " & (iC1RoundBet - iPRoundBet) & " chips to call.")
        Else    'Computer 1 normal bet
            iC1RoundBet = (iC1RoundBet + (CompChoice * iBlind))
            CheckLegal iC1RoundBet, iC1Chips, iPChips
            Alert("Computer 1 raises to " & iC1Roundbet & "; " & (iC1RoundBet - iPRoundBet) & " chips to call.")
        End If
    End Sub

    Sub PlayerCall()
        RaiseLock ""
        bC1Moved = False
        iPChips = (iPChips - (iC1RoundBet - iPRoundBet))    'Only reduce chips by re-raised value
        iPRoundBet = iC1RoundBet
        MoveToPot
        If iPChips = 0 or iC1Chips = 0 Then
            AllInRounds
        Else
            iRound = (iRound+1)
        End If
    End Sub

    Sub CompCall()
        If iPRoundBet = 0 Then
            Alert("Computer 1 checks.")
        Else
            Alert("Computer 1 calls your bet of " & iPRoundBet & " chips.")
        End If
        iC1Chips = (iC1Chips - (iPRoundBet - iC1RoundBet)) 'Only reduce chips by re-raised value
        iC1RoundBet = iPRoundBet
        MoveToPot
        bC1Moved = False
        If iC1Chips = 0 Or iPChips = 0 Then
            AllInRounds
        ElseIf iRound = 1 Then
            Flop
            iRound = (iRound+1)
        ElseIf iRound = 2 Then
            Turn
            iRound = (iRound+1)
        ElseIf iRound = 3 Then
            River
            iRound = (iRound+1)
        ElseIf iRound = 4 Then
            HandOver
        End If
    End Sub

    Sub CompFold()
        Fold iPChips, "Computer 1 folds the pot; You win."
        bC1Fold = True
    End Sub

    Sub Fold(WinnerChips, Desc)
        MoveToPot
        StartEndLock "Disabled", ""
        ShowBest Desc
        DistributePot WinnerChips, iPot
    End Sub

    Sub MoveToPot() 'Moves bets to chip pot and resets betting state for next round
        iPot = (iPot + iPRoundBet + iC1RoundBet)
        iPRoundBet = 0
        iC1Roundbet = 0
        iPChoice = 0
        ShowBet
    End Sub

    Sub ShowBet()   'Updates chip totals and player hand
        parPChips.innerHTML = "<Table><Tr><Th>Player</Th></Tr></Table><Table><Tr><Td>Chips: " & iPChips & "</Td></Tr><Tr><Td>" & sPBestHandDesc & "</Td></Tr></Table>"
        parC1Chips.innerHTML = "<Table><Tr><Th>Computer 1</Th></Tr></Table><Table><Tr><Td>Chips:</Td><Td>" & iC1Chips & "</Td></Tr></Table>"
        parResult.innerHTML = "<Table><Tr><Td>Current Pot:</Td><Td>" & iPot & "</Td></Tr></Table>"
    End Sub

    Sub DistributePot(WinnerChips, Chips)
        WinnerChips = (WinnerChips + Chips)
    End Sub