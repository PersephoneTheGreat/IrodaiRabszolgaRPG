Attribute VB_Name = "Actions"

Sub sewercide()

'Az összes action közül ezt írtam meg elõször FYI
    
    Dim ws As Worksheet
    Dim cell As Range
    Dim btn As Object
    Dim btnName As String
    
    'delete buttons
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'delete everything else too
    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    With ws.Range("A1:K16")
        .ClearContents
        .Merge
        .VerticalAlignment = xlVAlignTop
        .Value = "Ezt nem bírod tovább. Kiveted magad az irodaház ablakán." & vbCrLf & vbCrLf & _
        "Zuhansz... " & vbCrLf & vbCrLf & "Minden elsötétül."
    End With
    
    'add továb button
    ' Set the specific cell where you want to place the button
    Set cell = ActiveSheet.Range("E18:G19")
    
    ' Calculate the position and size based on the cell
    Dim btnWidth As Double, btnHeight As Double
    
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    
    ' Add a button to the worksheet at the specific cell
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    
    ' Set button properties
    With btn
        .Name = "btnRunMacro"
        .Caption = "Tovább"
    End With
    
    ' Assign the macro to the button
        btn.OnAction = "sewercide2"
    
End Sub

Sub Wait()
'várok action

    happening = "Kipihented magad"
    QuarterTime = QuarterTime + 1
    If Energy < 100 Then
        Energy = Energy + 1
    End If
    
    Call MainPage
    
End Sub

Sub Coffee()

    'Kávézunk
    If Booze >= 1 Then
        happening = "Megittál egy kávét"
        Anxiety = Anxiety + 0.1
        Booze = Booze - 1
        If Energy + 12 > 100 Then
            Energy = 100
        Else: Energy = Energy + 12
        End If
    Else: happening = "Nincs több kávéd"
    End If
    
    QuarterTime = QuarterTime + 1
    
    Call MainPage
End Sub

Sub EatXanax()

    'Xanaxozunk
    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    If Xanax >= 1 Then
        happening = "Bevettél egy Xanaxot"
        Xanax = Xanax - 1
        Anxiety = Anxiety - 0.1
    Else: happening = "Nincs több Xanaxod"
    End If
    
    
    Call MainPage
End Sub


Sub Curse()

    'káromkodunk
    Energy = Energy - 1
    Anxiety = Anxiety + 0.1
    QuarterTime = QuarterTime + 1
    Dim proclaim As Variant
    Dim userInput As String
    
    proclaim = Array("kiabálod", "ordítod", "köpöd oda", "sikítod", "ugatod", "morgod", "gondolod magadban")
    lowerBound = LBound(proclaim)
    upperBound = UBound(proclaim)
    Randomize
    randomIndex = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
    randomValue = proclaim(randomIndex)
    
    ' Prompt the user for input
    userInput = InputBox("Mit mondasz?")
    
    happening = userInput & " -- " & randomValue & "."
    
    'environment reacts
    Select Case Encounter
    
        Case "None"
            
        Case "Boss"
            happening = happening & " A fõnököd csúnyán néz, de egy darabig békén hagy."
            ifBoss = False
        
        Case "HR"
            happening = "A HR-es azonnal elküld egy Code of Conduct trainingre, és órákig ott rohadsz."
            Time = Time + 3
            
        Case "Stakeholder"
            happening = happening & " Ám a jenki egy szót sem ért magyarul, ezért csak mosolyog, bólogat, és még egy órán át beszél."
            Time = Time + 1
        
        Case "Boltos"
            happening = "Na, menj innen a fenébe, velem egy öltönyös így nem beszél! -- mondja a boltos, és kizavar."
            isManna = False
        
        Case "Email"
            happening = happening & " Amíg szitkozódsz, valaki más válaszol a levélre."
        
        Case "Projekt"
            happening = happening & " Amíg szitkozódsz, valamelyik kollégád megcsinálja a részed."
        
        Case "Jozsi"
            happening = happening & " Józsi felháborodva elhúzza a büdös seggét."
        
        Case "Jolan"
            happening = happening & " Jolán erre még cifrábbat válaszol vissza, teljesen belebolondulsz vodkaszagú szavaiba."
        
        Case "Betegszabi"
            happening = happening & " Látván felindultságodat Józsi önként jelentkezik, hogy helyettesítse Szabikát."
    
    End Select
    
    Encounter = "None"

    Call MainPage

End Sub

Sub buyBooze()

Dim Bandi As Double
Bandi = Rnd

Energy = Energy - 1
QuarterTime = QuarterTime + 1

    If isManna = True Then
        If Money >= 50 Then
            Money = Money - 50
                If Bandi <= 0.05 Then
                    happening = "Vettél kávét a Mannában, De amikor nem néztél oda, " & _
                    "az a fasz Bandi a reportingról ellopta."
                Else:
                    Booze = Booze + 1
                    happening = "Vettél kávét a Mannában."
                End If
        Else: happening = "Nincs elég pénzed kávéra."
        End If
    Else: happening = "Ha kávét akarsz venni, menj a Mannába!"
    End If
    
    Call MainPage

End Sub

Sub buyXanax()

Energy = Energy - 1
QuarterTime = QuarterTime + 1

    If isManna = True Then
        If Money > 100 Then
            Money = Money - 100
            Xanax = Xanax + 1
            happening = "Vettél egy szem Xanaxot a Mannás árufeltöltõ csávótól, aki stikában árulja."
        Else: happening = "Nincs elég pénzed Xanaxra."
        End If
    Else: happening = "Ha Xanaxot akarsz venni, menj a Mannába!"
    End If
    
    Call MainPage

End Sub

Sub proceed()

    Select Case Encounter
                
        Case "None", "Boltos"
            Energy = Energy - 1
            QuarterTime = QuarterTime + 1
            isManna = False
            Encounter = "None"
            
            Dim PossibleEncounters As Variant
            Dim EncounterProbabilities As Variant
            Dim cumulativeProbabilities As Variant
            Dim rndValue As Double
            Dim cumulativeSum As Double
            Dim i As Integer
            
            PossibleEncounters = Array("Neutral", _
                                        "Boss_Encounter", _
                                        "HR_Encounter", _
                                        "ProdTraining", _
                                        "Stakeholder_Encounter", _
                                        "Manna", _
                                        "Email", _
                                        "Projekt", _
                                        "Jozsi", _
                                        "Jolan", _
                                        "Betegszabi")
            EncounterProbabilities = Array(0.2, _
                                        0.02, _
                                        0.1, _
                                        0.1, _
                                        0.1, _
                                        0.13, _
                                        0.09, _
                                        0.09, _
                                        0.09, _
                                        0.09, _
                                        0.09)
                                        
             ' Calculate cumulative probabilities
                ReDim cumulativeProbabilities(UBound(EncounterProbabilities))
                cumulativeSum = 0
                For i = 0 To UBound(EncounterProbabilities)
                    cumulativeSum = cumulativeSum + EncounterProbabilities(i)
                    cumulativeProbabilities(i) = cumulativeSum
                Next i
            
                ' Generate a random value between 0 and 1
                rndValue = Rnd()
                
                For i = 0 To UBound(cumulativeProbabilities)
                    If rndValue <= cumulativeProbabilities(i) Then
                        Application.Run PossibleEncounters(i)
                        Exit For
                    End If
                Next i
                
            Case Else
            
            happening = "Azt hiszed, ilyen könnyen megúszod?"
            
            End Select
        
        
        Call MainPage
 
End Sub

Sub work()

QuarterTime = QuarterTime + 1

Dim productivity As Double
Dim WorkChance As Double

'munka

    Select Case Encounter
    
        Case "None"
            happening = "Nincs is meló, te hülye"
            Call MainPage
            
        Case "Boss"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 50
                    ifBoss = False
                    happening = "A fõnököd elégedetten konstatálja, hogy szorgalmasan dolgozol."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.1
                    happening = "Kidolgozod a beled, de ez senkit nem érdekel."
                End If
            Encounter = "None"
            Call MainPage
        
        Case "HR"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 50
                    ifBoss = False
                    happening = "A HR-es elégedetten konstatálja, hogy szorgalmasan dolgozol."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.1
                    happening = "Kidolgozod a beled, de ez senkit nem érdekel."
                End If
            Encounter = "None"
            Call MainPage
            
        Case "Stakeholder"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 100
                    happening = "A stakeholder nagyon örül annak, hogy ilyen szorgalmas vagy."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.1
                    happening = "Kidolgozod a beled, de ez senkit nem érdekel."
                End If
            Encounter = "None"
            Call MainPage
        
        Case "Boltos"
            happening = "-- Ne tedd az agyad, nem az irodában vagy. Na, veszel valamit? -- kérdi az eladó."
            Call MainPage
        
        Case "Email"
        
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 10
                    happening = "Válaszoltál az emailre"
                Else:
                    Energy = Energy - 10
                    happening = "Válaszolni akarsz, de a feladó email-fiókja megtelt, így nem jut el hozzá a válasz."
                End If
            Encounter = "None"
            Call MainPage
        
        Case "Projekt"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 50
                    happening = "Dolgozol egy kicsit a projekten."
                Else:
                    Energy = Energy - 10
                    happening = "Dolgozol egy kicsit a projekten, de senki le se szarja."
                End If
            Encounter = "None"
            Call MainPage
        
        Case "Jozsi"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 30
                    Money = Money + 100
                    happening = "Segítesz Józsinak a riporttal. A szaga elviselhetetlen."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.4
                    happening = "A riport sehogy nem akar formát ölteni. Józsi közben ott bûzölög melletted, ami iszonyatosan felbasz."
                End If
            Encounter = "None"
            Call MainPage
        
        Case "Jolan"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 500
                    happening = "Megcsinálod a taskot Jolán helyett. Vodkaszagú lehelletével rebeg neked hálát."
                Else:
                    Energy = Energy - 40
                    Anxiety = Anxiety + 0.2
                    happening = "Megcsinálod a taskot Jolán helyett, aki hálából az öledbe hányja másnapos gyomortartalmát. "
                End If
            Encounter = "None"
            Call MainPage
        
        Case "Betegszabi"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    happening = "Helyettesíted Szabikát, amíg beteg. Ezért nem jár plusz pénz."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.1
                    happening = "Megpróbálod helyettesíteni Szabikát, de ötleted sincs, mit csinált, mert nem vezeti a treckert."
                End If
            Encounter = "None"
            Call MainPage
        
    End Select
            

        
End Sub

Sub Slack()

QuarterTime = QuarterTime + 1
Dim Cunning As Double
Dim SlackChance As Double

'lógás

    Select Case Encounter
        
        Case "None"
            happening = "Elolvasod Facebookon az Irodai Rabszolga legújabb posztját"
            Call MainPage
            
        Case "Boss"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    Money = Money + 50
                    ifBoss = False
                    happening = "Bõszen püfölni kezded a billentyûzetet, mintha valami nagyon fontosat csinálnál, de csak shitposztolsz Redditre. A fõnököd azt hiszi, dolgozol."
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Elkezded leírni a paprikáskrumpli recepjét egy emailbe. Véletlenül rányomsz a küldés gombra. A levél a fõnöködnél landol."
                End If
            Encounter = "None"
            Call MainPage
        
        Case "HR"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Elkezded nyomkodni a mobilodat, mintha nem hallottál volna semmit, és kisétálsz a liftbõl."
                Else:
                    Anxiety = Anxiety + 0.1
                    happening = "Próbálsz úgy tenni, mintha fel se vennéd, de azért egész nap nyomaszt a dolog."
                End If
            Call MainPage
            
        Case "Stakeholder"
        
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "A meetingen videojátékokról kezdtek beszélgetni egy munkatársaddal, de mivel a jenki nem tud angolul, azt hiszi, saját projektbe kezdtetek. Örül, hogy ilyen produktívak vagytok."
                    ifStakeholder = False
                    Money = Money + 500
                Else:
                    Anxiety = Anxiety + 0.1
                    happening = "Bármivel próbálkozol, a jenki észreveszi, és téged kér meg, hogy prezentálj. Ez teljesen kimerít"
                    Energy = Energy - 30
                End If
            Call MainPage
        
        Case "Boltos"
            happening = "-- Lóghatsz itt, de akkor venned is kell valamit -- mondja a boltos, és beadod a derekad, veszel egy szendvicset."
            Money = Money - 150
            Call MainPage
        
        Case "Email"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Hasraütésbõl válaszolsz valami hülyeséget, de úgy tûnik, a feladónak tetszik."
                    Money = Money + 10
                Else:
                    Anxiety = Anxiety + 0.1
                    happening = "Hasraütésbõl válaszolsz valami hülyeséget, de a feladót nem lehet átverni, dühösen elküld a fenébe."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Projekt"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Úgy teszel, mintha csinálnál valamit. Senkinek nem tûnik fel, hogy nem."
                    Money = Money + 50
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Megpróbálos ellazázni a feladatot, de észreveszik. A team leaded behív elbeszélgetésre."
                    Energy = Energy - 20
                    Time = Time + 1
                End If
            Call MainPage
        
        Case "Jozsi"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Beleegyezel, de aztán szarsz a dologra. Végül Józsit basszák le, amiért nem készült el a feladat."
                    Money = Money + 30
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Beleegyezel, de aztán szarsz a dologra. Végül téged basznak le, amiért nincs kéz a feladat."
                End If
            Call MainPage
        
        Case "Jolan"
            'JOLÁN EGY ROMANCE-OLHATÓ KARAKTER AMÚGY
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Meggyõzöd Jolánt, hogy inkább menjetek be dugni az egyik meeting roomba. Vodkaíze van."
                    Money = Money + 30
                    Anxiety = 0.1
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Meggyõzöd Jolánt, hogy inkább menjetek be dugni az egyik meeting roomba. Rajtakapnak, és mindketten mehettek Code of Conduct tréningre."
                    Time = Time + 2
                End If
            Call MainPage
        
        Case "Betegszabi"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Helyettesíted Szabikát, reménykedve, hogy nem kap ma semmi feladatot. Így is lesz."
                    Money = Money + 30
                    Anxiety = 0.1
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Reménykedsz, hogy Szabikának ma nem lesz semmi bejövõ feladata. De kiderül, hogy a paraszt fél éve halogat egy projektet, ami a nyakadba szakad."
                    Time = Time + 2
                End If
            Call MainPage
        
    End Select
            

End Sub

Sub Escape()

Energy = Energy - 5
QuarterTime = QuarterTime + 1

Dim Flight As Double
Dim EscapeChance As Double

'menekülünk

    Select Case Encounter
    
        Case "None"
            happening = "Elmenekülsz. Mondjuk nem tudom, mi elõl. Paranoiás vagy?"
            Call MainPage
            
        Case "Boss"
        
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Kereket oldottál, és a fõnök szem elõl tévesztett. De azért megjegyezte, hogy nem dolgozol rendesen."
                    Encounter = "None"
                    Call MainPage
                Else:
                    Call Bossfight
                End If
        
        Case "HR"
        
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Úgy rohansz ki a liftbõl, mintha kergetnének."

                Else:
                    happening = "A lift kínosan lassan ér fel az emeletetekre. A HR-es közben végig megrovóan néz, és egyre erõsebb fingszagot áraszt."
                    Anxiety = Anxiety + 0.1
                End If
            Call MainPage
            
        Case "Stakeholder"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Valami hülye kifogással kimented magad."

                Else:
                    happening = "Megpróbálod kimenteni magad, de a jenki nem érti, és betuszkol a meetingterembe. Végig meg se kell szólalnod, de elmegy vele 3 órád."
                    Time = Time + 3
                End If
            Call MainPage
        
        Case "Boltos"
            Encounter = "None"
            isManna = False
            happeing = "Elmenekülsz a Mannából. Talán ennyire ijesztõen magasak az árak?"
            
            Call MainPage
        
        Case "Email"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Átküldöd az emailt valakinek, aki szerinted jobban ért hozzá."

                Else:
                    happening = "Tovább akarod küldeni az emailt, de viszajön, hogy mégis neked kell foglalkoznod vele."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Projekt"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Kitalálsz valami átlátszó kifogást, hogy most miért nem tudsz dolgozni rajta. Elhiszik."

                Else:
                    happening = "Kitalálsz valami átlátszó kifogást, hogy most miért nem tudsz dolgozni rajta. Nem hiszik el."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Jozsi"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "-- Nem, mert büdös vagy -- mondod. Józsi keres egy másik áldozatot."

                Else:
                    happening = "-- Nem, mert büdös vagy -- mondod. Józsi nevet, azt hiszi, hogy viccelsz, és átküldi a fájlokat."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Jolan"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Elmenekülsz Jolán vodkaszaga elõl."

                Else:
                    happening = "Megpróbálsz elmenekülni Jolán elõl, de utolér, leteper, és megfenyeget, hogy agyonver, ha nem csinálod meg helyette a taskot."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Betegszabi"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Kimented magad valami átlátszó kamuval."

                Else:
                    happening = "Azzal mented ki magad, hogy neked is sok dolgod van. De így a saját dolgaidat kell csinálnod, hogy elhiggyék."
                    Energy = Energy - 10
                End If
            Call MainPage
    
    End Select

End Sub
