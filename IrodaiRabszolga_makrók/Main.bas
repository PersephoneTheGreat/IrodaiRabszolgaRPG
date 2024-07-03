Attribute VB_Name = "Main"
Public Sub MainPage()

    Dim ws As Worksheet
    Dim cell As Range
    Dim btn As Object
    Dim btnName As String

    Set ws = ThisWorkbook.Worksheets("Irodai Rabszolga")
    ' Clear content in range A1:K21 on "Irodai Rabszolga" worksheet
    With ws.Range("A1:K21")
        .ClearContents
        .UnMerge
    End With
    
    'Delete button
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
    'Count time
    Debug.Print QuarterTime
    If QuarterTime >= 4 Then
        Time = Time + 1
        Debug.Print QuarterTime
        QuarterTime = 0
    End If
     
    
    'Check Time
    If Time > 17 Then
        Day = Day + 1
        Time = 9
        ifBoss = False
        ifStakeholder = False
        If Energy < 100 Then
            Energy = Energy + 10
        Else:
            Energy = 100
        End If
        happening = "Eltelt egy �jabb nap. �jszaka viszonylag kipihented magad, �gy �jult er�vel v�gsz bele a munk�ba!"
    End If
    
        'check day
        With ws.Range("H21:K21")
            .Merge
            .Value = "Eltelt napok sz�ma: " & Day
        End With
    
    'Write action
    With ws.Range("A1:K3")
        .Merge
        .Font.Size = 11
        .VerticalAlignment = xlVAlignTop
        .HorizontalAlignment = xlVAlignJustify
        .WrapText = True
        .Value = happening
    End With
    
    'Print stats
    With Range("A4:K7").Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Range("A4:K7").Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Range("A4:K7").Font.Size = 11
        
        
    With Range("A4:b4")
        .Merge
        .Value = "Energia :"
    End With
    Range("C4").Value = Energy & "%"
    
    With Range("A5:b5")
        .Merge
        .Value = "Idegess�g :"
    End With
    Range("C5").Value = Anxiety
    
    With Range("A6:b6")
        .Merge
        .Value = "P�nzed :"
    End With
    Range("C6").Value = Money & " Ft"
    
    With Range("A7:b7")
        .Merge
        .Value = "Xanax : "
    End With
    Range("C7").Value = Xanax & " db"
    
    With Range("f4:g4")
        .Merge
        .Value = "Pontos id� : "
    End With
    Range("h4").Value = Time & " �ra"
    
    With Range("f5:h5")
        .Merge
        .Font.Color = vbRed
    End With
    If ifStakeholder = True Then
        Range("f5:h5").Value = "MEETINGVESZ�LY!"
    End If
    
    With Range("f6:h6")
        .Merge
        .Font.Color = RGB(0, 176, 240)
    End With
    If ifBoss = True Then
        Range("f6:h6").Value = "Keres a f�n�k!"
    End If
    
    With Range("f7:g7")
        .Merge
        .Value = "K�v� :"
    End With
    Range("h7").Value = Booze & " db"
        
    'Add action buttons
    Dim leftPos As Double, topPos As Double
    Dim btnWidth As Double, btnHeight As Double
    
    '�szakra
    Set cell = ActiveSheet.Range("D9:H9")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Megyek �szakra"
    End With
        btn.OnAction = "proceed"
        
    'D�lre
    Set cell = ActiveSheet.Range("D10:H10")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Megyek d�lre"
    End With
        btn.OnAction = "proceed"
        
    'Keletre
    Set cell = ActiveSheet.Range("D11:H11")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Megyek keletre"
    End With
        btn.OnAction = "proceed"
        
    'Nyugatra
    Set cell = ActiveSheet.Range("D12:H12")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Megyek nyugatra"
    End With
        btn.OnAction = "proceed"
        
    'K�v�
    Set cell = ActiveSheet.Range("D13:H13")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Veszek k�v�t"
    End With
        btn.OnAction = "buyBooze"
        
    'Xanax
    Set cell = ActiveSheet.Range("D14:H14")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Veszek Xanaxot"
    End With
        btn.OnAction = "buyXanax"
        
    'Dolgozok
    Set cell = ActiveSheet.Range("D15:H15")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Dolgozok"
    End With
        btn.OnAction = "work"
        
    'L�gok
    Set cell = ActiveSheet.Range("D16:H16")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "L�gok"
    End With
        btn.OnAction = "Slack"
        
    'Menek�l�k
    Set cell = ActiveSheet.Range("D17:H17")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Menek�l�k"
    End With
        btn.OnAction = "Escape"
        
    'V�rok
    Set cell = ActiveSheet.Range("D18:H18")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "V�rok"
    End With
        btn.OnAction = "Wait"
        
    'Egy�b
    Set cell = ActiveSheet.Range("D19:H19")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Egy�b"
    End With
        btn.OnAction = "egyeb"
        
        
    'Check Anxiety levels
    If Anxiety < 0.1 Then
        Call ODhappens
    End If
    
    If Anxiety > 0.9 Then
        Call Breakdown
    End If
    
    'Check Energy levels
    If Energy < 1 Then
        Call burnout
    End If
    
End Sub

Sub egyeb()

 'Delete button
    Set btn = ActiveSheet.Buttons
    btn.Delete
    
 'Add egy�b buttons
    'K�v�zok
    Set cell = ActiveSheet.Range("D9:H9")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "K�v�zok"
    End With
        btn.OnAction = "Coffee"
    
    'Xanaxozok
    Set cell = ActiveSheet.Range("D10:H10")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Xanaxozok"
    End With
        btn.OnAction = "EatXanax"
        
    'K�romkodok
    Set cell = ActiveSheet.Range("D11:H11")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "K�romkodok"
    End With
        btn.OnAction = "Curse"
    
    '�ngyilkos leszek
    Set cell = ActiveSheet.Range("D12:H12")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "�ngyilkos leszek"
    End With
        btn.OnAction = "sewercide"
        
    'vissza (ez az eredeti CSOVES-ben nem volt, mindig hi�nyoltam)
    Set cell = ActiveSheet.Range("D19:H19")
    leftPos = cell.Left
    topPos = cell.Top
    btnWidth = cell.Width
    btnHeight = cell.Height
    Set btn = ActiveSheet.Buttons.Add(leftPos, topPos, btnWidth, btnHeight)
    With btn
        .Name = "btnRunMacro"
        .Caption = "Vissza"
    End With
        btn.OnAction = "MainPage"

End Sub
