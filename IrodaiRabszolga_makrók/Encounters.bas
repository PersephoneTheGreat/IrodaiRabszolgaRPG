Attribute VB_Name = "Encounters"
Sub Neutral()

    Dim neutralevent As Double
    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "None"
    
    neutralevent = Rnd
    If neutralevent < 0.5 Then
        happening = "Most �pp nincs semmi feladatod."
    Else
        happening = "Ez a meetingterem �res."
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
        happening = "A f�n�k gyanakodni kezdett, hogy t�l sokat l�gsz. Mostant�l jobban rajtad tartja a szem�t."
        Encounter = "Boss"
        Call MainPage
    End If

End Sub


Sub HR_Encounter()

    Energy = Energy - 10
    QuarterTime = QuarterTime + 1
    Anxiety = Anxiety + 0.1
    
    happening = "Belefutsz egy HR-esbe a liftben. Az megvet�en v�gigm�r, " & _
    "�s azt mondja: " & vbCrLf & " -- Ez a zenekaros fels� nem felel meg a dress code-nak!"
    Encounter = "HR"
    Call MainPage

End Sub

Sub ProdTraining()

    Energy = Energy - 5
    QuarterTime = QuarterTime + 1
    Time = 8
    Day = Day + 1
    Encounter = "None"
    
    happening = "Egy olyan meetingroomba nyitott�l be, ahol �ppen " & _
    "productivity-traininget tartottak. T�ged is ott fogtak eg�sz �jszak�ra. Az irod�ban �rt a hajnal."

    Call MainPage
    
End Sub

Sub Stakeholder_Encounter()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Anxiety = Anxiety + 0.1
    
    If ifStakeholder = True Then
        happening = "Odaj�n hozz�d az egyik jenki, aki azonnal beh�v az �j projektj�nek kick-off meetingj�re."
        Energy = Energy - 10
        Time = Time + 2
        Anxiety = Anxiety + 0.2
        Encounter = "Stakeholder"
    Else:
        ifStakeholder = True
        happening = "L�tod, ahogy az egyik jenki t�volr�l m�reget. Az a balsejtelmed t�mad, hogy be akar majd h�vni egy projektbe."
    End If
    
    Call MainPage

End Sub

Sub Manna()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    
    isManna = True
    happening = "Leugrott�l a Mann�ba. Itt az alkalom k�v�t vagy Xanaxot venni!"
    Encounter = "Boltos"
    
    Call MainPage

End Sub

Sub Email()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Email"
    happening = "Kapt�l egy emailt, amire v�laszolnod kell"
    
    Call MainPage
    

End Sub
Sub Projekt()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Projekt"
    happening = "K�zeledig a projekted hat�rideje, lassan neki k�ne �llni csin�lni valamit."
    
    Call MainPage

End Sub

Sub Jozsi()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Jozsi"
    happening = "J�zsi nem tud egyed�l �ssze-vlookup-ozni k�t riportot, �s megk�r, hogy seg�ts neki. (De j�l gondold meg, mit v�laszolsz; J�zsi nagyon b�d�s.)"
    
    Call MainPage

End Sub

Sub Jolan()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Jolan"
    happening = "Jol�n megint m�snaposan j�tt dolgozni, �s r�d akarja tolni a saj�t feladataidat. Cser�be megengedi, hogy megfogd a mell�t (csak a balt)."
    
    Call MainPage

End Sub

Sub Betegszabi()

    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    Encounter = "Betegszabi"
    happening = "Szabika lerobbant, �s t�ged adott meg helyettes�t�snek, am�g f�l nem �p�l. Valamit kezdeni k�ne a halmoz�d� ticketekkel."
    
    Call MainPage

End Sub



