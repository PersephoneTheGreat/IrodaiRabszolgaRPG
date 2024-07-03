Attribute VB_Name = "Actions"

Sub sewercide()

'Az �sszes action k�z�l ezt �rtam meg el�sz�r FYI
    
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
        .Value = "Ezt nem b�rod tov�bb. Kiveted magad az irodah�z ablak�n." & vbCrLf & vbCrLf & _
        "Zuhansz... " & vbCrLf & vbCrLf & "Minden els�t�t�l."
    End With
    
    'add tov�b button
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
        .Caption = "Tov�bb"
    End With
    
    ' Assign the macro to the button
        btn.OnAction = "sewercide2"
    
End Sub

Sub Wait()
'v�rok action

    happening = "Kipihented magad"
    QuarterTime = QuarterTime + 1
    If Energy < 100 Then
        Energy = Energy + 1
    End If
    
    Call MainPage
    
End Sub

Sub Coffee()

    'K�v�zunk
    If Booze >= 1 Then
        happening = "Megitt�l egy k�v�t"
        Anxiety = Anxiety + 0.1
        Booze = Booze - 1
        If Energy + 12 > 100 Then
            Energy = 100
        Else: Energy = Energy + 12
        End If
    Else: happening = "Nincs t�bb k�v�d"
    End If
    
    QuarterTime = QuarterTime + 1
    
    Call MainPage
End Sub

Sub EatXanax()

    'Xanaxozunk
    Energy = Energy - 1
    QuarterTime = QuarterTime + 1
    If Xanax >= 1 Then
        happening = "Bevett�l egy Xanaxot"
        Xanax = Xanax - 1
        Anxiety = Anxiety - 0.1
    Else: happening = "Nincs t�bb Xanaxod"
    End If
    
    
    Call MainPage
End Sub


Sub Curse()

    'k�romkodunk
    Energy = Energy - 1
    Anxiety = Anxiety + 0.1
    QuarterTime = QuarterTime + 1
    Dim proclaim As Variant
    Dim userInput As String
    
    proclaim = Array("kiab�lod", "ord�tod", "k�p�d oda", "sik�tod", "ugatod", "morgod", "gondolod magadban")
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
            happening = happening & " A f�n�k�d cs�ny�n n�z, de egy darabig b�k�n hagy."
            ifBoss = False
        
        Case "HR"
            happening = "A HR-es azonnal elk�ld egy Code of Conduct trainingre, �s �r�kig ott rohadsz."
            Time = Time + 3
            
        Case "Stakeholder"
            happening = happening & " �m a jenki egy sz�t sem �rt magyarul, ez�rt csak mosolyog, b�logat, �s m�g egy �r�n �t besz�l."
            Time = Time + 1
        
        Case "Boltos"
            happening = "Na, menj innen a fen�be, velem egy �lt�ny�s �gy nem besz�l! -- mondja a boltos, �s kizavar."
            isManna = False
        
        Case "Email"
            happening = happening & " Am�g szitkoz�dsz, valaki m�s v�laszol a lev�lre."
        
        Case "Projekt"
            happening = happening & " Am�g szitkoz�dsz, valamelyik koll�g�d megcsin�lja a r�szed."
        
        Case "Jozsi"
            happening = happening & " J�zsi felh�borodva elh�zza a b�d�s segg�t."
        
        Case "Jolan"
            happening = happening & " Jol�n erre m�g cifr�bbat v�laszol vissza, teljesen belebolondulsz vodkaszag� szavaiba."
        
        Case "Betegszabi"
            happening = happening & " L�tv�n felindults�godat J�zsi �nk�nt jelentkezik, hogy helyettes�tse Szabik�t."
    
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
                    happening = "Vett�l k�v�t a Mann�ban, De amikor nem n�zt�l oda, " & _
                    "az a fasz Bandi a reportingr�l ellopta."
                Else:
                    Booze = Booze + 1
                    happening = "Vett�l k�v�t a Mann�ban."
                End If
        Else: happening = "Nincs el�g p�nzed k�v�ra."
        End If
    Else: happening = "Ha k�v�t akarsz venni, menj a Mann�ba!"
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
            happening = "Vett�l egy szem Xanaxot a Mann�s �rufelt�lt� cs�v�t�l, aki stik�ban �rulja."
        Else: happening = "Nincs el�g p�nzed Xanaxra."
        End If
    Else: happening = "Ha Xanaxot akarsz venni, menj a Mann�ba!"
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
            
            happening = "Azt hiszed, ilyen k�nnyen meg�szod?"
            
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
            happening = "Nincs is mel�, te h�lye"
            Call MainPage
            
        Case "Boss"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 50
                    ifBoss = False
                    happening = "A f�n�k�d el�gedetten konstat�lja, hogy szorgalmasan dolgozol."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.1
                    happening = "Kidolgozod a beled, de ez senkit nem �rdekel."
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
                    happening = "A HR-es el�gedetten konstat�lja, hogy szorgalmasan dolgozol."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.1
                    happening = "Kidolgozod a beled, de ez senkit nem �rdekel."
                End If
            Encounter = "None"
            Call MainPage
            
        Case "Stakeholder"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 100
                    happening = "A stakeholder nagyon �r�l annak, hogy ilyen szorgalmas vagy."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.1
                    happening = "Kidolgozod a beled, de ez senkit nem �rdekel."
                End If
            Encounter = "None"
            Call MainPage
        
        Case "Boltos"
            happening = "-- Ne tedd az agyad, nem az irod�ban vagy. Na, veszel valamit? -- k�rdi az elad�."
            Call MainPage
        
        Case "Email"
        
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 10
                    happening = "V�laszolt�l az emailre"
                Else:
                    Energy = Energy - 10
                    happening = "V�laszolni akarsz, de a felad� email-fi�kja megtelt, �gy nem jut el hozz� a v�lasz."
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
                    happening = "Seg�tesz J�zsinak a riporttal. A szaga elviselhetetlen."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.4
                    happening = "A riport sehogy nem akar form�t �lteni. J�zsi k�zben ott b�z�l�g melletted, ami iszonyatosan felbasz."
                End If
            Encounter = "None"
            Call MainPage
        
        Case "Jolan"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    Money = Money + 500
                    happening = "Megcsin�lod a taskot Jol�n helyett. Vodkaszag� lehellet�vel rebeg neked h�l�t."
                Else:
                    Energy = Energy - 40
                    Anxiety = Anxiety + 0.2
                    happening = "Megcsin�lod a taskot Jol�n helyett, aki h�l�b�l az �ledbe h�nyja m�snapos gyomortartalm�t. "
                End If
            Encounter = "None"
            Call MainPage
        
        Case "Betegszabi"
            productivity = 1 - Anxiety
            WorkChance = Rnd
                If WorkChance <= productivity Then
                    Energy = Energy - 10
                    happening = "Helyettes�ted Szabik�t, am�g beteg. Ez�rt nem j�r plusz p�nz."
                Else:
                    Energy = Energy - 10
                    Anxiety = Anxiety + 0.1
                    happening = "Megpr�b�lod helyettes�teni Szabik�t, de �tleted sincs, mit csin�lt, mert nem vezeti a treckert."
                End If
            Encounter = "None"
            Call MainPage
        
    End Select
            

        
End Sub

Sub Slack()

QuarterTime = QuarterTime + 1
Dim Cunning As Double
Dim SlackChance As Double

'l�g�s

    Select Case Encounter
        
        Case "None"
            happening = "Elolvasod Facebookon az Irodai Rabszolga leg�jabb posztj�t"
            Call MainPage
            
        Case "Boss"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    Money = Money + 50
                    ifBoss = False
                    happening = "B�szen p�f�lni kezded a billenty�zetet, mintha valami nagyon fontosat csin�ln�l, de csak shitposztolsz Redditre. A f�n�k�d azt hiszi, dolgozol."
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Elkezded le�rni a paprik�skrumpli recepj�t egy emailbe. V�letlen�l r�nyomsz a k�ld�s gombra. A lev�l a f�n�k�dn�l landol."
                End If
            Encounter = "None"
            Call MainPage
        
        Case "HR"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Elkezded nyomkodni a mobilodat, mintha nem hallott�l volna semmit, �s kis�t�lsz a liftb�l."
                Else:
                    Anxiety = Anxiety + 0.1
                    happening = "Pr�b�lsz �gy tenni, mintha fel se venn�d, de az�rt eg�sz nap nyomaszt a dolog."
                End If
            Call MainPage
            
        Case "Stakeholder"
        
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "A meetingen videoj�t�kokr�l kezdtek besz�lgetni egy munkat�rsaddal, de mivel a jenki nem tud angolul, azt hiszi, saj�t projektbe kezdtetek. �r�l, hogy ilyen produkt�vak vagytok."
                    ifStakeholder = False
                    Money = Money + 500
                Else:
                    Anxiety = Anxiety + 0.1
                    happening = "B�rmivel pr�b�lkozol, a jenki �szreveszi, �s t�ged k�r meg, hogy prezent�lj. Ez teljesen kimer�t"
                    Energy = Energy - 30
                End If
            Call MainPage
        
        Case "Boltos"
            happening = "-- L�ghatsz itt, de akkor venned is kell valamit -- mondja a boltos, �s beadod a derekad, veszel egy szendvicset."
            Money = Money - 150
            Call MainPage
        
        Case "Email"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Hasra�t�sb�l v�laszolsz valami h�lyes�get, de �gy t�nik, a felad�nak tetszik."
                    Money = Money + 10
                Else:
                    Anxiety = Anxiety + 0.1
                    happening = "Hasra�t�sb�l v�laszolsz valami h�lyes�get, de a felad�t nem lehet �tverni, d�h�sen elk�ld a fen�be."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Projekt"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "�gy teszel, mintha csin�ln�l valamit. Senkinek nem t�nik fel, hogy nem."
                    Money = Money + 50
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Megpr�b�los ellaz�zni a feladatot, de �szreveszik. A team leaded beh�v elbesz�lget�sre."
                    Energy = Energy - 20
                    Time = Time + 1
                End If
            Call MainPage
        
        Case "Jozsi"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Beleegyezel, de azt�n szarsz a dologra. V�g�l J�zsit bassz�k le, ami�rt nem k�sz�lt el a feladat."
                    Money = Money + 30
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Beleegyezel, de azt�n szarsz a dologra. V�g�l t�ged basznak le, ami�rt nincs k�z a feladat."
                End If
            Call MainPage
        
        Case "Jolan"
            'JOL�N EGY ROMANCE-OLHAT� KARAKTER AM�GY
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Meggy�z�d Jol�nt, hogy ink�bb menjetek be dugni az egyik meeting roomba. Vodka�ze van."
                    Money = Money + 30
                    Anxiety = 0.1
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Meggy�z�d Jol�nt, hogy ink�bb menjetek be dugni az egyik meeting roomba. Rajtakapnak, �s mindketten mehettek Code of Conduct tr�ningre."
                    Time = Time + 2
                End If
            Call MainPage
        
        Case "Betegszabi"
            Encounter = "None"
            Cunning = 1 - (Anxiety * 2)
            SlackChance = Rnd
                If SlackChance <= Cunning Then
                    happening = "Helyettes�ted Szabik�t, rem�nykedve, hogy nem kap ma semmi feladatot. �gy is lesz."
                    Money = Money + 30
                    Anxiety = 0.1
                Else:
                    Anxiety = Anxiety + 0.2
                    happening = "Rem�nykedsz, hogy Szabik�nak ma nem lesz semmi bej�v� feladata. De kider�l, hogy a paraszt f�l �ve halogat egy projektet, ami a nyakadba szakad."
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

'menek�l�nk

    Select Case Encounter
    
        Case "None"
            happening = "Elmenek�lsz. Mondjuk nem tudom, mi el�l. Paranoi�s vagy?"
            Call MainPage
            
        Case "Boss"
        
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Kereket oldott�l, �s a f�n�k szem el�l t�vesztett. De az�rt megjegyezte, hogy nem dolgozol rendesen."
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
                    happening = "�gy rohansz ki a liftb�l, mintha kergetn�nek."

                Else:
                    happening = "A lift k�nosan lassan �r fel az emeletetekre. A HR-es k�zben v�gig megrov�an n�z, �s egyre er�sebb fingszagot �raszt."
                    Anxiety = Anxiety + 0.1
                End If
            Call MainPage
            
        Case "Stakeholder"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Valami h�lye kifog�ssal kimented magad."

                Else:
                    happening = "Megpr�b�lod kimenteni magad, de a jenki nem �rti, �s betuszkol a meetingterembe. V�gig meg se kell sz�lalnod, de elmegy vele 3 �r�d."
                    Time = Time + 3
                End If
            Call MainPage
        
        Case "Boltos"
            Encounter = "None"
            isManna = False
            happeing = "Elmenek�lsz a Mann�b�l. Tal�n ennyire ijeszt�en magasak az �rak?"
            
            Call MainPage
        
        Case "Email"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "�tk�ld�d az emailt valakinek, aki szerinted jobban �rt hozz�."

                Else:
                    happening = "Tov�bb akarod k�ldeni az emailt, de viszaj�n, hogy m�gis neked kell foglalkoznod vele."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Projekt"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Kital�lsz valami �tl�tsz� kifog�st, hogy most mi�rt nem tudsz dolgozni rajta. Elhiszik."

                Else:
                    happening = "Kital�lsz valami �tl�tsz� kifog�st, hogy most mi�rt nem tudsz dolgozni rajta. Nem hiszik el."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Jozsi"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "-- Nem, mert b�d�s vagy -- mondod. J�zsi keres egy m�sik �ldozatot."

                Else:
                    happening = "-- Nem, mert b�d�s vagy -- mondod. J�zsi nevet, azt hiszi, hogy viccelsz, �s �tk�ldi a f�jlokat."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Jolan"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Elmenek�lsz Jol�n vodkaszaga el�l."

                Else:
                    happening = "Megpr�b�lsz elmenek�lni Jol�n el�l, de utol�r, leteper, �s megfenyeget, hogy agyonver, ha nem csin�lod meg helyette a taskot."
                    Energy = Energy - 10
                End If
            Call MainPage
        
        Case "Betegszabi"
            Encounter = "None"
            Flight = 1.3 - Anxiety
            EscapeChance = Rnd
                If EscapeChance <= Flight Then
                    happening = "Kimented magad valami �tl�tsz� kamuval."

                Else:
                    happening = "Azzal mented ki magad, hogy neked is sok dolgod van. De �gy a saj�t dolgaidat kell csin�lnod, hogy elhiggy�k."
                    Energy = Energy - 10
                End If
            Call MainPage
    
    End Select

End Sub
