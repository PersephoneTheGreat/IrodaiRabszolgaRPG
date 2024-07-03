Attribute VB_Name = "Encounters"
Sub Neutral()

    Dim neutralevent As Double
    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "None"
    
    neutralevent = Rnd
    If neutralevent < 0.5 Then
        happening = "Most épp nincs semmi feladatod."
    Else
        happening = "Ez a meetingterem üres."
    End If
    
    Call MainPage
    
End Sub

Sub Boss_Encounter()
    
    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Anxiety = Anxiety + 0.1
    
    If ifBoss = True Then
        Call Bossfight
    Else:
        ifBoss = True
        happening = "A fõnök gyanakodni kezdett, hogy túl sokat lógsz. Mostantól jobban rajtad tartja a szemét."
        Encounter = "Boss"
        Call MainPage
    End If

End Sub


Sub HR_Encounter()

    Energy = Energy - 10
    QuarterTime = QuarterTime + 1
    Anxiety = Anxiety + 0.1
    
    happening = "Belefutsz egy HR-esbe a liftben. Az megvetõen végigmér, " & _
    "és azt mondja: " & vbCrLf & " -- Ez a zenekaros felsõ nem felel meg a dress code-nak!"
    Encounter = "HR"
    Call MainPage

End Sub

Sub ProdTraining()

    Energy = Energy - 5
    QuarterTime = QuarterTime + 1
    Time = 8
    Day = Day + 1
    Encounter = "None"
    
    happening = "Egy olyan meetingroomba nyitottál be, ahol éppen " & _
    "productivity-traininget tartottak. Téged is ott fogtak egész éjszakára. Az irodában ért a hajnal."

    Call MainPage
    
End Sub

Sub Stakeholder_Encounter()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Anxiety = Anxiety + 0.1
    
    If ifStakeholder = True Then
        happening = "Odajön hozzád az egyik jenki, aki azonnal behív az új projektjének kick-off meetingjére."
        Energy = Energy - 10
        Time = Time + 2
        Anxiety = Anxiety + 0.2
        Encounter = "Stakeholder"
    Else:
        ifStakeholder = True
        happening = "Látod, ahogy az egyik jenki távolról méreget. Az a balsejtelmed támad, hogy be akar majd hívni egy projektbe."
    End If
    
    Call MainPage

End Sub

Sub Manna()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    
    isManna = True
    happening = "Leugrottál a Mannába. Itt az alkalom kávét vagy Xanaxot venni!"
    Encounter = "Boltos"
    
    Call MainPage

End Sub

Sub Email()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Email"
    happening = "Kaptál egy emailt, amire válaszolnod kell"
    
    Call MainPage
    

End Sub
Sub Projekt()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Projekt"
    happening = "Közeledig a projekted határideje, lassan neki kéne állni csinálni valamit."
    
    Call MainPage

End Sub

Sub Jozsi()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Jozsi"
    happening = "Józsi nem tud egyedül össze-vlookup-ozni két riportot, és megkér, hogy segíts neki. (De jól gondold meg, mit válaszolsz; Józsi nagyon büdös.)"
    
    Call MainPage

End Sub

Sub Jolan()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Jolan"
    happening = "Jolán megint másnaposan jött dolgozni, és rád akarja tolni a saját feladataidat. Cserébe megengedi, hogy megfogd a mellét (csak a balt)."
    
    Call MainPage

End Sub

Sub Betegszabi()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Betegszabi"
    happening = "Szabika lerobbant, és téged adott meg helyettesítésnek, amíg föl nem épül. Valamit kezdeni kéne a halmozódó ticketekkel."
    
    Call MainPage

End Sub



