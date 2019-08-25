Attribute VB_Name = "Module1"
'#####################################################################
'## Project: Monopoly Simulation
'##
'##
'## Script purpose: to evaluate the value of various monopoly properties
'##                 using a Markov type simulation
'##
'## Date:    July 24, 2014
'## Author:  Steven Smith
'#####################################################################

  Dim blnDouble As Boolean
  Dim datDate As Date
  Dim datTimeStart As Date
  Dim datTimeEnd As Date
  Dim intActivePlayer As Integer
  Dim intDiceRoll As Integer
  Dim intDicePrior As Integer
  Dim intMove As Integer
  Dim intMovesPerTest As Integer
  Dim intPlayers As Integer
  Dim intRow As Integer
  Dim intTests As Integer
  Dim intTestID As Integer
  Dim sngElapsedTime As Single
  Dim sngLevelOfSignificance As Single
  Dim strEvent As String
  Dim strTokens(1 To 9) As String
  Dim varPlayerData(1 To 8, 1 To 6) As Variant ' creates a two dimentional Player information array for up to 8 players with 6 properties (1: Token; 2: Cash Balance; 3: location; 4: roll-count; 5: status; 6: open)

Sub Simulation()
Attribute Simulation.VB_ProcData.VB_Invoke_Func = " \n14"
    
  '20140723.  This is a monopoly simulation intended to perform tests of hit frequencies for each property on a monopoly board.
  'Enter test in the log

  Sheets("Test Log").Select
  Range("B3").Select
  'Establish Values
      intTests = InputBox("How many Tests would you like to run? (between 1 and 1000)", "Test Parameters", 50)
      If intTests < 2 Then intTests = 2
      If intTests > 300 Then intTests = 300
      intMovesPerTest = InputBox("How many Moves per Test would you like to run? (between 1 and 1000)", "Test Parameters", 50)
      If intMovesPerTest < 30 Then intMovesPerTest = 30
      If intMovesPerTest > 300 Then intMovesPerTest = 300
      sngLevelOfSignificance = InputBox("What Level of Significance Would you like to use?", "Test Parameters", 0.01)
      intTestID = ActiveCell.Value + 1
      intRow = intTestID + 8
      datDate = Year(Now()) & "/" & Month(Now()) & "/" & Day(Now())
      datTimeStart = Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now())
      intPlayers = InputBox("How many players? (between 1 and 8)", "Player Count", 4)
      If intPlayers < 1 Then intPlayers = 4
        If intPlayers > 8 Then intPlayers = 4
      MsgBox ("Player count has been set at " & intPlayers)
    'Write Values in header
      Range("B3").Select
      ActiveCell.FormulaR1C1 = intTestID
      Range("B4").Select
      ActiveCell.FormulaR1C1 = datDate
      Range("B5").Select
      ActiveCell.FormulaR1C1 = datTimeStart
      Range("B6").Select
      ActiveCell.FormulaR1C1 = intPlayers
      
    'Locate Current Record Line
      Range("A8").Select
      For i = 1 To intTestID
        ActiveCell.Offset(1, 0).Range("A1").Select
      Next
    'Write Current Entry
      ActiveCell.FormulaR1C1 = intTestID
      ActiveCell.Offset(0, 1).Range("A1").Select
      ActiveCell.FormulaR1C1 = datDate
      ActiveCell.Offset(0, 1).Range("A1").Select
      ActiveCell.FormulaR1C1 = datTime
      ActiveCell.Offset(0, 1).Range("A1").Select
      ActiveCell.FormulaR1C1 = intPlayers
      ActiveCell.Offset(0, 1).Range("A1").Select
      ActiveCell.FormulaR1C1 = intMovesPerTest
      ActiveCell.Offset(0, 1).Range("A1").Select
      ActiveCell.FormulaR1C1 = intTests
      ActiveCell.Offset(0, 1).Range("A1").Select
      ActiveCell.FormulaR1C1 = sngLevelOfSignificance
    'Prepare game
      Sheets("Current Test").Select
        Range("B3:ALM2002").Select
        Selection.ClearContents
        Range("B3").Select
        ActiveCell.Offset(0, Z).Range("A1").Select
        intActivePlayer = 1
        intDicePrior = 0
        For i = 1 To intPlayers
          Call DiceRoll(intDiceRoll, blnDouble)
          If intDiceRoll > intDicePrior Then
            intActivePlayer = i
            intDicePrior = intDiceRoll
          End If
          MsgBox ("player " & i & " rolled a " & intDiceRoll)
        Next i
        MsgBox ("player " & intActivePlayer & " rolled the highest and will go first")
        ' Populate Token Values
          strTokens(1) = "Hat"
          strTokens(2) = "Shoe"
          strTokens(3) = "Battleship"
          strTokens(4) = "Dog"
          strTokens(5) = "Cat"
          strTokens(6) = "Car"
          strTokens(7) = "Iron"
          strTokens(8) = "Wheelbarrow"
        'Populate Player Data Array
          For i = 1 To intPlayers
            varPlayerData(i, 1) = strTokens(i)  '1: Token
            varPlayerData(i, 2) = 1500          '2: Cash Balance
            varPlayerData(i, 3) = 0             '3: Location
            varPlayerData(i, 4) = 0             '4: Roll-Count
            varPlayerData(i, 5) = "Free"        '5: Status (Free, Inmate, Bankrupt)
            varPlayerData(i, 6) = 0             '6: Open
          Next
        'Begin Game
          Sheets("Current Test").Select
          Range("B3").Select
          For Z = 1 To intTests
            Range("A3").Select
            ActiveCell.Offset(0, Z).Range("A1").Select
            For intMove = 1 To intMovesPerTest
              strEvent = ""
              If blnDouble = False Then
                varPlayerData(intActivePlayer, 4) = 0
              End If
              Call DiceRoll(intDiceRoll, blnDouble)
              If varPlayerData(intActivePlayer, 5) = "Inmate" Then
                'for this test, we'll presume the fee is paid.
                'Decide to try and get out
                Randomize
                
                    If blnDouble = True Then
                    varPlayerData(intActivePlayer, 5) = "Free"
                End If
              Else
                If varPlayerData(intActivePlayer, 3) + intDiceRoll > 39 Then
                  varPlayerData(intActivePlayer, 3) = varPlayerData(intActivePlayer, 3) + intDiceRoll - 40
                  varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 200
                Else
                  varPlayerData(intActivePlayer, 3) = varPlayerData(intActivePlayer, 3) + intDiceRoll
                End If
              'test for three doubles in a row
              If blnDouble = True Then
                If varPlayerData(intActivePlayer, 4) = 2 Then
                  varPlayerData(intActivePlayer, 3) = 10
                  varPlayerData(intActivePlayer, 5) = "Inmate"
                End If
              End If
              varPlayerData(intActivePlayer, 4) = varPlayerData(intActivePlayer, 4) + 1
              'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
              'ActiveCell.Offset(0, 1).Range("A1").Select
              'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
              'ActiveCell.Offset(0, 1).Range("A1").Select
              'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
              'ActiveCell.Offset(0, 1).Range("A1").Select
              'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
              'ActiveCell.Offset(0, 1).Range("A1").Select
              'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
              'ActiveCell.Offset(0, 1).Range("A1").Select
              'ActiveCell.FormulaR1C1 = intDiceRoll
              'ActiveCell.Offset(0, 1).Range("A1").Select
              'ActiveCell.FormulaR1C1 = blnDouble
              'ActiveCell.Offset(0, 1).Range("A1").Select
              'ActiveCell.FormulaR1C1 = strEvent
              'ActiveCell.Offset(1, -7).Range("A1").Select
              ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
              ActiveCell.Offset(1, 0).Range("A1").Select
              Select Case varPlayerData(intActivePlayer, 3)
                Case 0
                Case 1
                Case 2 ' Community Chest
                  Call CommunityChest(varPlayerData)
                Case 3
                Case 4
                Case 5
                Case 6
                Case 7 ' Chance
                  Call Chance(varPlayerData)
                Case 8
                Case 9
                Case 10
                Case 11
                Case 12
                Case 13
                Case 14
                Case 15
                Case 16
                Case 17 ' Community Chest
                  Call CommunityChest(varPlayerData)
                Case 18
                Case 19
                Case 20
                Case 21
                Case 22 ' Chance
                  Call Chance(varPlayerData)
                Case 23
                Case 24
                Case 25
                Case 26
                Case 27
                Case 28
                Case 29
                Case 30 ' Go To Jail
                  Call GoToJail(varPlayerData)
                Case 31
                Case 32
                Case 33 ' Community Chest
                  Call CommunityChest(varPlayerData)
                Case 34
                Case 35
                Case 36 ' Chance
                  Call Chance(varPlayerData)
                Case 37
                Case 38
                Case 39
              End Select
            'Set up next player
              If blnDouble = False Then
                If intPlayers > intActivePlayer Then
                  intActivePlayer = intActivePlayer + 1
                Else
                  intActivePlayer = 1
                End If
              ElseIf varPlayerData(intActivePlayer, 5) = "Inmate" Then
                If intPlayers > intActivePlayer Then
                  intActivePlayer = intActivePlayer + 1
                Else
                  intActivePlayer = 1
                End If
              End If
            Next
          Next
    Sheets("Group Analysis").Select
    Range("D2").Select
    ActiveCell.FormulaR1C1 = sngLevelOfSignificance
    Range("D3").Select
    
    Sheets("Group Analysis").Select
    Range("C7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Test Log").Select
    Range("A8").Select
      For i = 1 To intTestID
        ActiveCell.Offset(1, 0).Range("A1").Select
      Next
      ActiveCell.Offset(0, 7).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Range("A8").Select

    Sheets("Group Analysis").Select
    Range("E7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Test Log").Select
    Range("A8").Select
      For i = 1 To intTestID
        ActiveCell.Offset(1, 0).Range("A1").Select
      Next
      ActiveCell.Offset(0, 20).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    datTimeEnd = Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now())
    ActiveCell.Offset(0, 13).Range("A1").Select
    ActiveCell.FormulaR1C1 = datTimeEnd
    Range("A8").Select
    
    
  End Sub

Function DiceRoll(intDiceRoll, blnDouble)

  Dim intDice1 As Integer
  Dim intDice2 As Integer
    Randomize
    intDice1 = Round(Rnd() * 6 + 0.5, 0)
    intDice2 = Round(Rnd() * 6 + 0.5, 0)
    intDiceRoll = intDice1 + intDice2
    If intDice1 = intDice2 Then
      blnDouble = True
    Else
      blnDouble = False
    End If
End Function

Function CurrentPlayer(intActivePlayer)
  
End Function

Sub DiceRollTest()
  Dim intDice1 As Integer
  Dim intDice2 As Integer
    
  For k = 1 To 500
    For i = 1 To 500
      intDice1 = Round(Rnd() * 6 + 0.5, 0)
      ActiveCell.FormulaR1C1 = intDice1
      ActiveCell.Offset(1, 0).Range("A1").Select
    Next
    ActiveCell.Offset(-500, 1).Range("A1").Select
  Next
End Sub

Sub Test()
  Dim strTokens(1 To 9) As String
  
  strTokens(1) = "Hat"
  strTokens(2) = "Shoe"
  strTokens(3) = "Battleship"
  strTokens(4) = "Dog"
  strTokens(5) = "Cat"
  strTokens(6) = "Car"
  strTokens(7) = "Iron"
  strTokens(8) = "Wheelbarrow"
  For i = 1 To 8
    ActiveCell.FormulaR1C1 = i
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = strTokens(i)
    ActiveCell.Offset(1, -1).Range("A1").Select
  Next
End Sub

Function Card(intChance)
  intChance = Round(Rnd() * 16 + 0.5, 0)
End Function

Function CommunityChest(varPlayerData)
Call Card(intChance)
  Select Case intChance
    
    Case 1 'advance to Go
      strEvent = "Community Chest: Advance to Go"
      varPlayerData(intActivePlayer, 3) = 0
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 200
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = "D"
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = blnDouble
      'ActiveCell.Offset(-1, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(2, -7).Range("A1").Select
      ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      ActiveCell.Offset(1, 0).Range("A1").Select
    Case 2 ' Bank error in your favor - collect $75
      strEvent = "Community Chest: Bank error in your favor"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 75
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 3 ' Doctor's fees - pay $50.00
      strEvent = "Community Chest: Doctors Fees"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) - 50
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 4 ' Get out of jail free
      strEvent = "Community Chest: Get out of jail free"
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 5 ' Go to Jail
      strEvent = "Community Chest: Go to Jail"
      varPlayerData(intActivePlayer, 3) = 10
      varPlayerData(intActivePlayer, 5) = "Inmate"
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = "D"
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = blnDouble
      'ActiveCell.Offset(-1, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(2, -7).Range("A1").Select
      ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      ActiveCell.Offset(1, 0).Range("A1").Select

    Case 6 ' It is your birthday Collect $10 from each player
      strEvent = "Community Chest: It's your birthday"
      For i = 1 To intPlayers
        If intActivePlayer = i Then
          varPlayerData(i, 2) = varPlayerData(i, 2) + (intPlayers - 1) * 10
        Else
          varPlayerData(i, 2) = varPlayerData(i, 2) - 10
        End If
      Next
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 7 ' Grand opera night - collect $50 from every player
      strEvent = "Community Chest: Grand Opera Night"
      For i = 1 To intPlayers
        If intActivePlayer = i Then
          varPlayerData(i, 2) = varPlayerData(i, 2) + (intPlayers - 1) * 50
        Else
          varPlayerData(i, 2) = varPlayerData(i, 2) - 50
        End If
      Next
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
      
    Case 8 ' Income tax refund collect $20
      strEvent = "Community Chest: Income tax refund"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 20
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 9 ' Life Insurance Matures - collect $100
      strEvent = "Community Chest: Life Insurance Matures"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 100
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 10 'Pay Hospital Fees of $100
      strEvent = "Community Chest: Pay Hospital Fees  "
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) - 100
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 11 ' Owe School Fees of $50
      strEvent = "Community Chest: School Fees"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) - 50
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 12 ' Receive $25 Consultancy Fee
      strEvent = "Community Chest: Receive Consultancy Fee"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 25
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 13 ' You are assessed for street repairs - $40 per house, $115 per hotel
      strEvent = "Community Chest: Street Repairs"
   '   ActiveCell.Offset(-1, 7).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = strEvent
   '   ActiveCell.Offset(1, -7).Range("A1").Select
        
    Case 14 ' You have won second prize in a beauty contest - collect $10
      strEvent = "Community Chest: Second prise in a beauty contest"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 10
   '   ActiveCell.Offset(-1, 7).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = strEvent
   '   ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 15 ' You inherit $100
      strEvent = "Community Chest: You inherit $100"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 100
   '   ActiveCell.Offset(-1, 7).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = strEvent
   '   ActiveCell.Offset(1, -7).Range("A1").Select
    
    Case 16 ' From sale of stock you get $50
      strEvent = "Community Chest: Sale of Stock"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 50
   '   ActiveCell.Offset(-1, 7).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = strEvent
   '   ActiveCell.Offset(1, -7).Range("A1").Select
  
  End Select
End Function

Function Chance(varPlayerData)
  Call Card(intChance)
  Select Case intChance
    
    Case 1 'advance to Go
      strEvent = "Chance: Advance to Go"
      varPlayerData(intActivePlayer, 3) = 0
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 200
    '  ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
    '  ActiveCell.Offset(0, 1).Range("A1").Select
    '  ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
    '  ActiveCell.Offset(0, 1).Range("A1").Select
    '  ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
    '  ActiveCell.Offset(0, 1).Range("A1").Select
    '  ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
    '  ActiveCell.Offset(0, 1).Range("A1").Select
    '  ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
    '  ActiveCell.Offset(0, 1).Range("A1").Select
    '  ActiveCell.FormulaR1C1 = "D"
    '  ActiveCell.Offset(0, 1).Range("A1").Select
    '  ActiveCell.FormulaR1C1 = blnDouble
    '  ActiveCell.Offset(-1, 1).Range("A1").Select
    '  ActiveCell.FormulaR1C1 = strEvent
    '  ActiveCell.Offset(2, -7).Range("A1").Select
   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
    
    Case 2 ' Advance to Illinois Ave.
      strEvent = "Chance: Advance to Illinois Ave"
      If varPlayerData(intActivePlayer, 3) > 24 Then
        varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 200
      End If
      varPlayerData(intActivePlayer, 3) = 24
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = "D"
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = blnDouble
      'ActiveCell.Offset(-1, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(2, -7).Range("A1").Select
     ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
          
    Case 3 ' Advance token to nearest Utility.  If unowned, you may buy if from the bank.  If owned, throw dice and pay owner a total ten times the amount thrown.
      strEvent = "Chance: Advance to Nearest Utility"
      If varPlayerData(intActivePlayer, 3) < 12 Then
        varPlayerData(intActivePlayer, 3) = 12
      ElseIf varPlayerData(intActivePlayer, 3) < 28 Then
        varPlayerData(intActivePlayer, 3) = 28
      Else
        varPlayerData(intActivePlayer, 3) = 12
        varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 200
      End If
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = "D"
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = blnDouble
      'ActiveCell.Offset(-1, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(2, -7).Range("A1").Select
   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
          
    Case 4 ' Advance token to the nearest Railroad and pay owner twice the rental to which he/she is otherwise entitled.  If railroad is unowned, you may buy it from the Bank.
      strEvent = "Chance: Advance to nearest Railroad (1)"
      If varPlayerData(intActivePlayer, 3) < 5 Then
        varPlayerData(intActivePlayer, 3) = 5
      ElseIf varPlayerData(intActivePlayer, 3) < 15 Then
        varPlayerData(intActivePlayer, 3) = 15
      ElseIf varPlayerData(intActivePlayer, 3) < 25 Then
        varPlayerData(intActivePlayer, 3) = 25
      ElseIf varPlayerData(intActivePlayer, 3) < 35 Then
        varPlayerData(intActivePlayer, 3) = 35
      Else
        varPlayerData(intActivePlayer, 3) = 5
        varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 200
      End If
   '   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
   '   ActiveCell.Offset(0, 1).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
   '   ActiveCell.Offset(0, 1).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
   '   ActiveCell.Offset(0, 1).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
   '   ActiveCell.Offset(0, 1).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
   '   ActiveCell.Offset(0, 1).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = "D"
   '   ActiveCell.Offset(0, 1).Range("A1").Select
   '   ActiveCell.FormulaR1C1 = blnDouble
   '   ActiveCell.Offset(-1, 1).Range("A1").Select
    '  ActiveCell.FormulaR1C1 = strEvent
   '   ActiveCell.Offset(2, -7).Range("A1").Select
   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
          
    Case 5 ' Second of two:  Advance token to the nearest Railroad and pay owner twice the rental to which he/she is otherwise entitled.  If railroad is unowned, you may buy it from the Bank.
      strEvent = "Chance: Advance to nearest Railroad (2)"
      If varPlayerData(intActivePlayer, 3) < 5 Then
        varPlayerData(intActivePlayer, 3) = 5
      ElseIf varPlayerData(intActivePlayer, 3) < 15 Then
        varPlayerData(intActivePlayer, 3) = 15
      ElseIf varPlayerData(intActivePlayer, 3) < 25 Then
        varPlayerData(intActivePlayer, 3) = 25
      ElseIf varPlayerData(intActivePlayer, 3) < 35 Then
        varPlayerData(intActivePlayer, 3) = 35
      Else
        varPlayerData(intActivePlayer, 3) = 5
        varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 200
      End If
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = "D"
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = blnDouble
      'ActiveCell.Offset(-1, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(2, -7).Range("A1").Select
   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
          
    Case 6 ' Advance to St. Charles Place - if you pass Go, collect $200.
      strEvent = "Chance: Advance to St. Charles Place"
      If varPlayerData(intActivePlayer, 3) > 11 Then
        varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 200
      End If
      varPlayerData(intActivePlayer, 3) = 11
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = "D"
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = blnDouble
      'ActiveCell.Offset(-1, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(2, -7).Range("A1").Select
     ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
          
    Case 7 ' Bank pays you dividend of $50.00
      strEvent = "Chance: Bank pays you dividend of $50"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 50
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
     
    Case 8 ' Get out of jail free
      strEvent = "Chance: Get out of Jail free"
     ' ActiveCell.Offset(-1, 7).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = strEvent
     ' ActiveCell.Offset(1, -7).Range("A1").Select
     '
    Case 9 ' Go back three spaces.
      strEvent = "Chance: Go back three spaces"
      If varPlayerData(intActivePlayer, 3) < 3 Then
        varPlayerData(intActivePlayer, 3) = varPlayerData(intActivePlayer, 3) - 3 + 40
      Else
        varPlayerData(intActivePlayer, 3) = varPlayerData(intActivePlayer, 3) - 3
      End If
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = "D"
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = blnDouble
      'ActiveCell.Offset(-1, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(2, -7).Range("A1").Select
   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
          
    Case 10 ' Go to Jail
      strEvent = "Chance: Go to Jail"
      varPlayerData(intActivePlayer, 3) = 10
      varPlayerData(intActivePlayer, 5) = "Inmate"
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = "D"
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = blnDouble
      'ActiveCell.Offset(-1, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(2, -7).Range("A1").Select
   ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
          
    Case 11 ' Make general repairs on all your property - for each house pay $25, $100 per hotel
      strEvent = "Chance: Make general repairs"
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
      
    Case 12 'Pay poor tax of $15
      strEvent = "Chance: Pay poor tax of $15"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) - 15
     ' ActiveCell.Offset(-1, 7).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = strEvent
     ' ActiveCell.Offset(1, -7).Range("A1").Select
      
    Case 13 ' Take a trip on the Reading Railroad - if you pass Go collect $200
      strEvent = "Chance: Take a trip on the Reading Railroad"
      If varPlayerData(intActivePlayer, 3) > 5 Then
        varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 200
      End If
      varPlayerData(intActivePlayer, 3) = 5
     ' ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
     ' ActiveCell.Offset(0, 1).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
     ' ActiveCell.Offset(0, 1).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
     ' ActiveCell.Offset(0, 1).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
     ' ActiveCell.Offset(0, 1).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
     ' ActiveCell.Offset(0, 1).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = "D"
     ' ActiveCell.Offset(0, 1).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = blnDouble
     ' ActiveCell.Offset(-1, 1).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = strEvent
     ' ActiveCell.Offset(2, -7).Range("A1").Select
     ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
          
    Case 14 ' Take a walk on the Boardwalk - advance token to Boardwalk
      strEvent = "Chance: Take a walk on the Boardwalk"
      varPlayerData(intActivePlayer, 3) = 39
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = "D"
      'ActiveCell.Offset(0, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = blnDouble
      'ActiveCell.Offset(-1, 1).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(2, -7).Range("A1").Select
     ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select
          
    Case 15 ' You have been elected chairman of the board - pay each player $50
       strEvent = "Chance: Elected chairman of the board"
       For i = 1 To intPlayers
        If intActivePlayer = i Then
          varPlayerData(i, 2) = varPlayerData(i, 2) - (intPlayers - 1) * 50
        Else
          varPlayerData(i, 2) = varPlayerData(i, 2) + 50
        End If
      Next
      'ActiveCell.Offset(-1, 7).Range("A1").Select
      'ActiveCell.FormulaR1C1 = strEvent
      'ActiveCell.Offset(1, -7).Range("A1").Select
      
    Case 16 ' Your building loan matures - collect $150
      strEvent = "Chance: Your building loan matures"
      varPlayerData(intActivePlayer, 2) = varPlayerData(intActivePlayer, 2) + 150
     ' ActiveCell.Offset(-1, 7).Range("A1").Select
     ' ActiveCell.FormulaR1C1 = strEvent
     ' ActiveCell.Offset(1, -7).Range("A1").Select
      
  End Select
End Function

Function GoToJail(varPlayerData)
  strEvent = "Go To Jail"
  varPlayerData(intActivePlayer, 3) = 10
  varPlayerData(intActivePlayer, 5) = "Inmate"
  'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 1)
  'ActiveCell.Offset(0, 1).Range("A1").Select
  'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 2)
  'ActiveCell.Offset(0, 1).Range("A1").Select
  'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
  'ActiveCell.Offset(0, 1).Range("A1").Select
  'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 4)
  'ActiveCell.Offset(0, 1).Range("A1").Select
  'ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 5)
  'ActiveCell.Offset(0, 1).Range("A1").Select
  'ActiveCell.FormulaR1C1 = "D"
  'ActiveCell.Offset(0, 1).Range("A1").Select
  'ActiveCell.FormulaR1C1 = blnDouble
  'ActiveCell.Offset(-1, 1).Range("A1").Select
  '    ActiveCell.FormulaR1C1 = strEvent
  '    ActiveCell.Offset(2, -7).Range("A1").Select
  ActiveCell.FormulaR1C1 = varPlayerData(intActivePlayer, 3)
          ActiveCell.Offset(1, 0).Range("A1").Select         
End Function
