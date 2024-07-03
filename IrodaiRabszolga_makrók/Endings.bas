Attribute VB_Name = "Endings"
Sub sewercide2()
    
        'delete buttons
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'delete everything else too
    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    With ws.Range("A1:K16")
        .ClearContents
        .Merge
        .VerticalAlignment = xlVAlignTop
    End With
    
    ' Set the text in the range A1:K16
    Range("A1:K16").Value = "A telefonod csörgésére ébredsz, reggel 7 óra van. " & _
    "Megint az irodával álmodtál." & vbCrLf & vbCrLf & _
    "Szerencsére nem sok mindenre emlékszel. De nem is sok idõd van merengeni rajta, " & _
    "hiszen lassan indulnod kell dolgozni. " & vbCrLf & vbCrLf & _
    "Vár a Váci utca!"
    
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
        btn.OnAction = "new_game"

End Sub

Sub ODhappens()

    'delete buttons
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'delete everything else too
    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    With ws.Range("A1:K16")
        .ClearContents
        .Merge
        .VerticalAlignment = xlVAlignTop
        .Value = "Túl sok Xanaxot vettél be" & vbCrLf & vbCrLf & "Minden elsötétül."
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

Sub Breakdown()

'Alapvetõen akkor jön elõ, ha az idegesség túl magasra megy,
'de lesz valamennyi random valószínûsége is, hogy kipörgeted, mert miért ne

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
        .Value = "Addig idegesítetted magad, hogy végül idegösszeomlást kaptál. " & _
        "Mentõ vitt el a Váci útról." & vbCrLf & vbCrLf & "Szerencsére néhány hónap zárt osztály" & _
        " után már elég jól vagy ahhoz, hogy visszatérj régi pozíciódba."
    End With
    
    'add továb button
    'Set the specific cell where you want to place the button
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
        btn.OnAction = "New_game"

End Sub

Sub burnout()

 'delete buttons
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'delete everything else too
    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    With ws.Range("A1:K16")
        .ClearContents
        .Merge
        .VerticalAlignment = xlVAlignTop
        .Value = "Minden erõd elfogyott, teljesen kiégtél." & vbCrLf & vbCrLf & _
        "Szerencsére az ottalvós pszichiátria és a motivációs trénerek " & _
        "új életet és motivációt leheltek " & _
        "csoffadt lelkedbe, és hamarosan visszatérhetsz az irodába."
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
        btn.OnAction = "New_game"


End Sub

Sub Bossfight()
'Elkap a fõnök ending
'delete buttons
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'delete everything else too
    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    With ws.Range("A1:K16")
        .ClearContents
        .Merge
        .VerticalAlignment = xlVAlignTop
        .Value = "A fõnök megelégelte, hogy ennyit lógsz, és behívat az irodájába." & vbCrLf & vbCrLf & _
                 "-- Nézd, én csak a javadat akarom. Tudom, hogy téged csak a többi sóher rángatott bele. " & _
                 "De én segíthetek neked produktívabbá válni -- mondja, és elküld tréningre" & vbCrLf & vbCrLf & _
                 "Addig ülsz a feladatmenedzsment, stresszkezelés, és compliance tréningeken, amíg belezöldülsz." & vbCrLf & vbCrLf & _
                 "Ezután újult erõvel, rengeteg soft skillel felszerelve látsz újra neki a munkának."
        
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
        btn.OnAction = "New_game"

End Sub

Sub SecretEndingScreen1()

    'Ez egy olyan advanced játék, hogy még secret endingje is van, gyá
    'delete buttons
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'delete everything else too
    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    With ws.Range("A1:K16")
        .ClearContents
        .Merge
        .VerticalAlignment = xlVAlignTop
        .Value = "Végre összeszedted a bátorságodat, és felmondtál. Régi álmodat követve saját céget alapítottál"
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
        btn.OnAction = "SecretEndingScreen2"

End Sub

Sub SecretEndingScreen2()

    'delete buttons
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'delete everything else too
    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    With ws.Range("A1:K16")
        .ClearContents
        .Merge
        .VerticalAlignment = xlVAlignTop
        .Value = "Néhány hónappal késõbb beüt egy újabb gazdasági válság, a céged csõdbe megy " & _
        "Szerencsére a régi munkahelyed szívesen visszavesz ugyanabba a pozícióba." & vbCrLf & vbCrLf & _
        "Újra vár a Váci utca!"
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
        btn.OnAction = "new_game"

End Sub

'01000001 00100000 01101101 01110101 01101100 01110100 01101001 01100011 01100101 01100111 00101100 00100000
'01100001 01101000 01101111 01101100 00100000 01100100 01101111 01101100 01100111 01101111 01111010 01101111
'01101101 00101100 00100000 01100101 01101100 01101011 01110101 01101100 01100100 01101111 01110100 01110100
'00100000 01100101 01100111 01111001 00100000 01101101 01100001 01101011 01110010 01101111 00100000 01110100
'01110010 01100101 01101110 01101001 01101110 01100111 01110010 01100101 00101100 00100000 01100001 01101101
'01101001 01110100 00100000 01110101 01100111 01111001 00100000 01111010 01100001 01110010 01110100 01100001
'01101011 00101100 00100000 01101000 01101111 01100111 01111001 00100000 01100111 01111001 01100001 01101011
'01101111 01110010 01101111 01101100 01101010 01110101 01101110 01101011 00100000 01101101 01100001 01100111
'01110101 01101110 01101011 01110100 01101111 01101100 00101100 00100000 01110011 01100001 01101010 01100001
'01110100 00100000 01110000 01110010 01101111 01101010 01100101 01101011 01110100 01100101 01101011 01100101
'01101110 00101110 00100000 01010011 01101111 00100000 01001001 00100000 01100100 01101001 01100100 00100000
'01100101 01111000 01100001 01100011 01110100 01101100 01111001 00100000 01110100 01101000 01100001 01110100
'00101110 00100000 01010011 01101111 01101000 01100001 00100000 01101110 01100101 01101101 00100000 01110110
'01101111 01101100 01110100 00100000 01100001 01100110 01100110 01101001 01101110 01101001 01110100 01100001
'01110011 01101111 01101101 00100000 01100001 01111010 00100000 01101001 01101110 01100110 01101111 01110010
'01101101 01100001 01110100 01101001 01101011 01100001 01101000 01101111 01111010 00101100 00100000 01100010
'01101111 01101100 01100011 01110011 01100101 01110011 01111010 00100000 01110110 01100001 01100111 01111001
'01101111 01101011 00101100 00100000 00101000 01100010 01100001 01110010 00100000 01100001 01101100 01110100
'01100001 01101100 01100001 01101110 01101111 01110011 01100010 01100001 01101110 00100000 01100101 01100111
'01111001 01110011 01111010 01100101 01110010 00100000 01101101 01100101 01100111 01101110 01111001 01100101
'01110010 01110100 01100101 01101101 00100000 01100001 00100000 01101101 01100101 01100111 01111001 01100101
'01101001 00100000 01000011 01101111 01101101 01101100 01101111 01100111 01101111 00100000 01110110 01100101
'01110010 01110011 01100101 01101110 01111001 01110100 00101100 00100000 01100100 01100101 00100000 01100011
'01110011 01100001 01101011 00100000 01100001 01111010 01100101 01110010 01110100 00101100 00100000 01101101
'01100101 01110010 01110100 00100000 01110010 01100001 01101010 01110100 01100001 01101101 00100000 01101011
'01101001 01110110 01110101 01101100 00100000 01101110 01100101 01101101 00100000 01101001 01101110 01100100
'01110101 01101100 01110100 00100000 01110011 01100101 01101110 01101011 01101001 00101001 00101100 00100000
'01110011 01111010 01101111 01110110 01100001 01101100 00100000 01101000 01100001 00100000 01100001 00100000
'01101011 01101111 01100100 00100000 01110011 01111010 01100001 01110010 00101100 00100000 01100001 01111010
'00100000 01100001 00100000 01101000 01101111 01111010 01111010 01100001 01101110 01100101 01101101 01100101
'01110010 01110100 01100101 01110011 01100101 01101101 00100000 01101101 01101001 01100001 01110100 01110100
'00100000 01110110 01100001 01101110 00101110 00100000 01000010 01100001 01110010 00100000 01110101 01100111
'01111001 00100000 01101000 01100001 01101100 01101100 01101111 01110100 01110100 01100001 01101101 00101100
'00100000 01100001 01111010 00100000 01100001 00100000 01101101 01110101 01101100 01110100 01101001 01101110
'01100001 01101100 00100000 01110110 01100101 01111010 01100101 01110100 01101111 01101001 00100000 01101011
'01110110 01100001 01101100 01101001 01110100 01100001 01110011 00101110 00100000 01000001 00100000 01110000
'01110010 01101111 01101010 01100101 01101011 01110100 01100101 01110100 00100000 01100001 01111010 00100000
'01001001 01110010 01101111 01100100 01100001 01101001 00100000 01010010 01100001 01100010 01110011 01111010
'01101111 01101100 01100111 01100001 00100000 01100110 01100001 01100011 01100101 01100010 01101111 01101111
'01101011 00101101 01101111 01101100 01100100 01100001 01101100 01101111 01101110 00100000 01101011 01101001
'01110110 01110101 01101100 00100000 01100001 00100000 01101100 01100101 01100111 01100101 01101110 01100100
'01100001 01110011 00100000 00100111 00111000 00111001 00101101 01100101 01110011 00100000 01010011 01011010
'01000001 01001011 01000001 01000100 01010100 00100000 01000011 01010011 01001111 01010110 01000101 01010011
'00100000 01101110 01100101 01110110 01110101 00100000 01101010 01100001 01110100 01100101 01101011 00100000
'01101001 01101000 01101100 01100101 01110100 01110100 01100101 00101110 00100000 01000101 01111010 01110100
'00100000 01100001 00100000 01101110 01100101 01110100 01100101 01101110 00100000 01110100 01100101 01110010
'01101010 01100101 01101110 01100111 01101111 00100000 01101000 01101001 01110010 01100101 01101011 01101011
'01100101 01101100 00100000 01100101 01101100 01101100 01100101 01101110 01110100 01100101 01110100 01100010
'01100101 01101110 00100000 01101110 01100101 01101101 00100000 01010100 01101111 01101101 01100011 01100001
'01110100 00100000 01101011 01100101 01110011 01111010 01101001 01110100 01100101 01110100 01110100 01100101
'00101100 00100000 01101000 01100001 01101110 01100101 01101101 00100000 01010011 01111010 01111001 00100000
'01011010 01101111 01101100 01110100 01100001 01101110 00101100 00100000 01100001 01101011 01101001 00100000
'01101101 01100001 00100000 01100101 01100111 01111001 00100000 01101101 01110101 01101100 01110100 01101001
'01100011 01100101 01100111 01101110 01100101 01101100 00100000 01001001 01010100 00100000 01010011 01100101
'01100011 01110101 01110010 01101001 01110100 01111001 00100000 01110100 01100101 01100001 01101101 00100000
'01101100 01100101 01100001 01100100 00100000 01100001 00100000 01101100 01101001 01101110 01101011 01100101
'01100100 01101001 01101110 01101010 01100101 00100000 01110011 01111010 01100101 01110010 01101001 01101110
'01110100 00101110 00100000 01000001 01111010 00100000 01101001 01110011 00100000 01100101 01100111 01111001
'00100000 01110011 01111010 01100101 01110000 00100000 01100101 01101100 01100101 01110100 01110100 01101111
'01110010 01110100 01100101 01101110 01100101 01110100 00100000 01101100 01100101 01101000 01100101 01110100
'101110
