    Sub Buttons(Change1, Change2)
        Dim i
        btnFold.disabled = Change1
        For i = 0 to 5
            document.getElementById("btnBet" & i).disabled = Change1
        Next
        btnStart.disabled = Change2
    End Sub

    Sub CalcBlind()
        iBlind = ((Int(iHandNum/10)+1)*50)
        iAnte = ((Int(iHandNum/20)-1)*25)
        If iAnte < 0 Then
            iAnte = 0
        End If
    End Sub

    Sub CalcBet(GreatScore, VeryHighScore, HighScore, MedScore, LowScore)
        If iRound = 0 Then
            iPRoundBet = iAnte
'           If (iHandNum / 2) = Int(iHandNum / 2) Then
                iPRoundBet = iPRoundBet + iBlind
                CheckLegal
                iPChips = iPChips - iPRoundBet
                CompMove GreatScore, VeryHighScore, HighScore, MedScore, LowScore, (iBlind/2)
'           Else
'               iPChips = iPChips - (iBlind/2)
'               iPRoundBet = (iBlind/2)
'                CompMove GreatScore, VeryHighScore, HighScore, MedScore, LowScore, iBlind
'           End If
        Else If iChoice < 5 Then
                iPRoundBet = iChoice * iBlind
                CheckLegal
                CompMove GreatScore, VeryHighScore, HighScore, MedScore, LowScore, 0
            ElseIf iChoice = 5 Then
                iPRoundBet = iPChips
                CheckLegal
                CompMove GreatScore, VeryHighScore, HighScore, MedScore, LowScore, 0
            End If
        End If
    End Sub

    Sub CheckLegal()
        If iPRoundBet > iPChips Then
            iPRoundBet = iPChips
        ElseIf iPRoundBet > iC1Chips Then
            iPRoundBet = iC1Chips
        End If
        iPChips = iPChips - iPRoundBet
    End Sub

    Sub CompMove(GreatScore, VeryHighScore, HighScore, MedScore, LowScore, PayBlind)
        If iC1BestScore => GreatScore And iChoice = 5 Or iC1BestScore => VeryHighScore And iChoice <= 4 Or iC1BestScore => HighScore And iChoice <= 3 _
        Or iC1BestScore => MedScore And iChoice <= 2 Or iC1BestScore => LowScore And iChoice <= 1 Or iC1BestScore => 0 And iChoice <= 0 Then
            CompCall
            ShowBet
        Else
            CompFold PayBlind
        End If
    End Sub

    Sub CompCall()
        iC1RoundBet = iPRoundBet
        iC1Chips = iC1Chips - iC1RoundBet
        MoveToPot
    End Sub

    Sub CompFold(PayBlind)
        If iRound = 0 Then
            iC1Chips = iC1Chips - (PayBlind + iAnte)
            iC1RoundBet = iAnte + PayBlind
        End If
        MoveToPot
        Fold iPChips, "Computer 1 folds the pot; You win."
        bC1Fold = True
    End Sub

    Sub Fold(WinnerChips, Desc)
        Buttons "Disabled", ""
        ShowBet
        ShowBest Desc
        DistributePot WinnerChips, iPot
        iRound = 0
        iChoice = 0
    End Sub

    Sub MoveToPot()
        iPot = iPot + iPRoundBet + iC1RoundBet
        iPRoundBet = 0
        iC1Roundbet = 0
    End Sub

    Sub ShowBet()
        parPChips.innerHTML = "<Table><Tr><Th>Player</Th></Tr></Table><Table><Tr><Td>Chips:</Td><Td>" & iPChips & "</Td></Tr></Table>" & iPBestScore
        parC1Chips.innerHTML = "<Table><Tr><Th>Computer 1</Th></Tr></Table><Table><Tr><Td>Chips:</Td><Td>" & iC1Chips & "</Td></Tr></Table>" & iC1BestScore
        parResult.innerHTML = "<Table><Tr><Td>Current Pot:</Td><Td>" & iPot & "</Td></Tr></Table>"
    End Sub

    Sub DistributePot(WinnerChips, Chips)
        WinnerChips = WinnerChips + Chips
        Chips = 0
    End Sub