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
    Range("A1:K16").Value = "A telefonod cs�rg�s�re �bredsz, reggel 7 �ra van. " & _
    "Megint az irod�val �lmodt�l." & vbCrLf & vbCrLf & _
    "Szerencs�re nem sok mindenre eml�kszel. De nem is sok id�d van merengeni rajta, " & _
    "hiszen lassan indulnod kell dolgozni. " & vbCrLf & vbCrLf & _
    "V�r a V�ci utca!"
    
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
        .Value = "T�l sok Xanaxot vett�l be" & vbCrLf & vbCrLf & "Minden els�t�t�l."
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

Sub Breakdown()

'Alapvet�en akkor j�n el�, ha az idegess�g t�l magasra megy,
'de lesz valamennyi random val�sz�n�s�ge is, hogy kip�rgeted, mert mi�rt ne

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
        .Value = "Addig ideges�tetted magad, hogy v�g�l ideg�sszeoml�st kapt�l. " & _
        "Ment� vitt el a V�ci �tr�l." & vbCrLf & vbCrLf & "Szerencs�re n�h�ny h�nap z�rt oszt�ly" & _
        " ut�n m�r el�g j�l vagy ahhoz, hogy visszat�rj r�gi poz�ci�dba."
    End With
    
    'add tov�b button
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
        .Caption = "Tov�bb"
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
        .Value = "Minden er�d elfogyott, teljesen ki�gt�l." & vbCrLf & vbCrLf & _
        "Szerencs�re az ottalv�s pszichi�tria �s a motiv�ci�s tr�nerek " & _
        "�j �letet �s motiv�ci�t leheltek " & _
        "csoffadt lelkedbe, �s hamarosan visszat�rhetsz az irod�ba."
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
        btn.OnAction = "New_game"


End Sub

Sub Bossfight()
'Elkap a f�n�k ending
'delete buttons
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'delete everything else too
    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    With ws.Range("A1:K16")
        .ClearContents
        .Merge
        .VerticalAlignment = xlVAlignTop
        .Value = "A f�n�k megel�gelte, hogy ennyit l�gsz, �s beh�vat az irod�j�ba." & vbCrLf & vbCrLf & _
                 "-- N�zd, �n csak a javadat akarom. Tudom, hogy t�ged csak a t�bbi s�her r�ngatott bele. " & _
                 "De �n seg�thetek neked produkt�vabb� v�lni -- mondja, �s elk�ld tr�ningre" & vbCrLf & vbCrLf & _
                 "Addig �lsz a feladatmenedzsment, stresszkezel�s, �s compliance tr�ningeken, am�g belez�ld�lsz." & vbCrLf & vbCrLf & _
                 "Ezut�n �jult er�vel, rengeteg soft skillel felszerelve l�tsz �jra neki a munk�nak."
        
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
        btn.OnAction = "New_game"

End Sub

Sub SecretEndingScreen1()

    'Ez egy olyan advanced j�t�k, hogy m�g secret endingje is van, gy�
    'delete buttons
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'delete everything else too
    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    With ws.Range("A1:K16")
        .ClearContents
        .Merge
        .VerticalAlignment = xlVAlignTop
        .Value = "V�gre �sszeszedted a b�tors�godat, �s felmondt�l. R�gi �lmodat k�vetve saj�t c�get alap�tott�l"
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
        .Value = "N�h�ny h�nappal k�s�bb be�t egy �jabb gazdas�gi v�ls�g, a c�ged cs�dbe megy " & _
        "Szerencs�re a r�gi munkahelyed sz�vesen visszavesz ugyanabba a poz�ci�ba." & vbCrLf & vbCrLf & _
        "�jra v�r a V�ci utca!"
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
