Attribute VB_Name = "Intro"
Option Explicit

' Declare global variables
Public Energy As Double ' Percent (0-100)
Public Anxiety As Double ' Number between 0 and 1
Public Money As Integer ' Integer
Public Xanax As Integer ' Integer
Public QuarterTime As Integer 'r�szid�
Public Time As Integer ' 4 r�szid� = �ra
Public Booze As Integer ' K�v�
Public ifStakeholder As Boolean 'true or false
Public Attacker As Variant 'Ki t�mad �ppen?
Public Encounter As Variant 'Milyen projekttel vagy mobbal n�zel szembe?
Public ifBoss As Boolean 'true or false
Public happening As String 'esem�nyle�r�s
Public isManna As Boolean 'A Mann�ban vagy-e
Public Day As Integer 'h�ny nap telt el




Sub irodai_rabszolga_intro()

    Dim ws As Worksheet
    Dim btn As Object
    Dim btnName As String
    Dim sheetCount As Integer
    Dim sheetExists As Boolean
    Dim newSheet As Worksheet
    Dim cell As Range
    Dim targetRange As Range
    Dim shape As shape
    Dim leftPos As Double, topPos As Double, rightPos As Double, bottomPos As Double
    
    sheetExists = False
    
 ' Check if sheet named "Irodai Rabszolga" exists
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Irodai Rabszolga" Then
            sheetExists = True
            Set newSheet = ws ' Assign existing sheet to newSheet variable
            Exit For
        End If
    Next ws
    
    ' If sheet does not exist, create it
    If Not sheetExists Then
        Set newSheet = ThisWorkbook.Sheets.Add
        newSheet.Name = "Irodai Rabszolga"
    Else
        newSheet.Activate
    End If
    

    ' Format the range A1:K21
    With newSheet.Range("A1:K21")
        .UnMerge
        .ClearContents
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Interior.Color = 0
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "OCR A Extended"
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    
    'Delete buttons if any
    ' Define the target range where buttons should be deleted
    Set targetRange = ws.Range("A1:K21")
    
    ' Get the positions of the target range
    leftPos = targetRange.Left
    topPos = targetRange.Top
    rightPos = targetRange.Left + targetRange.Width
    bottomPos = targetRange.Top + targetRange.Height
    
    ' Loop through all shapes on the worksheet
    For Each shape In ws.Shapes
        ' Check if the shape is within the target range
        If shape.Type = msoFormControl Then ' Only consider form control shapes (buttons)
            If shape.Left >= leftPos And shape.Top >= topPos And _
               shape.Left + shape.Width <= rightPos And _
               shape.Top + shape.Height <= bottomPos Then
                shape.Delete
            End If
        End If
    Next shape
    
    ' Merge range B6:J7 and set its value
    With newSheet.Range("B3:J7")
        .Merge
        .Font.Size = 28
        .Value = "PersephoneProduction"
    End With
    
    With newSheet.Range("B8:J9")
        .Merge
        .Font.Size = 20
        .Value = "Presents"
    End With

    With newSheet.Range("B10:J16")
        .Merge
        .Font.Size = 36
        .Value = "Irodai Rabszolga '24"
    End With
    
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
        .Caption = "�j j�t�k"
    End With
    
    ' Assign the macro to the button
        btn.OnAction = "New_game"

End Sub

Sub New_game()

    Dim ws As Worksheet
    Dim cell As Range
    Dim btn As Object
    Dim btnName As String
    Randomize
    
    With ActiveSheet.Range("A1:K16")
        .ClearContents
        .UnMerge
        .Merge
        .Font.Size = 16
        .HorizontalAlignment = xlJustify
        .VerticalAlignment = xlJustify
        .WrapText = True
        .Value = "Te egy irodai rabszolga vagy a V�ci �ton. " & _
        "Az a c�lod, hogy min�l t�bb l�v�t szerezz. " & _
        "A Mann�ban vehetsz k�v�t �s Xanaxot. " & _
        "J�, ha vigy�zol magadra, mert olyan h�rek terjengenek, " & _
        "hogy jenki stakeholderek j�ttek Pestre. " & _
        "De az�rt nem kell mindj�rt berezelni, " & _
        "csak tegy�l �gy, mintha dolgozn�l. " & _
        "Az idegess�ged befoly�solhajta a menek�l�s �s a munka hat�konys�g�t. " & _
        "Ha gyenge vagy, igy�l egy k�v�t, ha idegesnek �rzed magad, Xanaxozz egy j�t."
    End With
    
    'set game stats to starting value
    Energy = 99
    Anxiety = 0.2
    Money = 1000
    Xanax = 4
    Time = 15
    Booze = 1
    Encounter = "None"
    ifStakeholder = False
    ifBoss = False
    isManna = True
    Day = 0
    happening = "A f�n�k�d gyanakodva n�z, de azt�n od�bb�ll. " & _
        "Ezt meg�sztad, nem vette �szre, hogy porn�t n�zel a c�ges g�pen."
    
     ' Set the specific cell where you want to place the button
    Set cell = ActiveSheet.Range("E18:G19")
    
    ' Calculate the position and size based on the cell
    Dim leftPos As Double, topPos As Double
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
        btn.OnAction = "MainPage"

End Sub

