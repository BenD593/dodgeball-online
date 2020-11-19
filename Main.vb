Imports System.Text.RegularExpressions
Imports System
Imports System.IO
Imports System.Reflection
Module Main
    ' Public Variables for Dice Image '
    Public myAssembly As Assembly = Assembly.GetExecutingAssembly()
    Public steDice1 As Stream = myAssembly.GetManifestResourceStream("dodgeball_online.1_D6.png")
    Public steDice2 As Stream = myAssembly.GetManifestResourceStream("dodgeball_online.2_D6.png")
    Public steDice3 As Stream = myAssembly.GetManifestResourceStream("dodgeball_online.3_D6.png")
    Public steDice4 As Stream = myAssembly.GetManifestResourceStream("dodgeball_online.4_D6.png")
    Public steDice5 As Stream = myAssembly.GetManifestResourceStream("dodgeball_online.5_D6.png")
    Public steDice6 As Stream = myAssembly.GetManifestResourceStream("dodgeball_online.6_D6.png")

    ' Public Variables for Game '
    Public strBlueA1, strBlueA2, strBlueA3, strBlueA4, strBlueA5, strBlueB1, strBlueB2, strBlueB3, strBlueB4, strBlueB5, strBlueC1, strBlueC2, strBlueC3, strBlueC4, strBlueC5, strBlueD1, strBlueD2, strBlueD3, strBlueD4, strBlueD5, strBlueE1, strBlueE2, strBlueE3, strBlueE4, strBlueE5 As String
    Public strGreenA1, strGreenA2, strGreenA3, strGreenA4, strGreenA5, strGreenB1, strGreenB2, strGreenB3, strGreenB4, strGreenB5, strGreenC1, strGreenC2, strGreenC3, strGreenC4, strGreenC5, strGreenD1, strGreenD2, strGreenD3, strGreenD4, strGreenD5, strGreenE1, strGreenE2, strGreenE3, strGreenE4, strGreenE5 As String
    Public intRound, intPlayerRound, intThrowNumber, intMissDirection, intGreenScore, intBlueScore As Integer
    Public strTarget, strResult As String
    Public random As New Random()

    '#################################'
    '# Sub Procedure to Ouput Result #'
    '#################################'
    Sub subResult()

        intThrowNumber = random.Next(1, 7)
        intMissDirection = random.Next(1, 9)

        subDieImage(intThrowNumber)
        subMissArrow(intMissDirection)

        If intThrowNumber = 1 Or intThrowNumber = 2 Then
            frmGameActive.lblDice.Text = intThrowNumber & Environment.NewLine() & "Miss by: " & fucThrowCalculation(intThrowNumber)
            frmGameActive.lblMissDirection.Text = intMissDirection & Environment.NewLine() & "Miss direction: " & fucMissCalculation(intMissDirection)
            frmGameActive.lblResult.Text = "You missed by: " & fucThrowCalculation(intThrowNumber) & " In the direction of: " & fucMissCalculation(intMissDirection) & Environment.NewLine() & "So: " & fucHitCheck(intThrowNumber, intMissDirection)
        ElseIf intThrowNumber = 3 Or intThrowNumber = 4 Then
            frmGameActive.lblDice.Text = intThrowNumber & Environment.NewLine() & "Miss by: " & fucThrowCalculation(intThrowNumber)
            frmGameActive.lblMissDirection.Text = intMissDirection & Environment.NewLine() & "Miss direction" & fucMissCalculation(intMissDirection)
            frmGameActive.lblResult.Text = "You missed by: " & fucThrowCalculation(intThrowNumber) & " In the direction of: " & fucMissCalculation(intMissDirection) & Environment.NewLine() & "So: " & fucHitCheck(intThrowNumber, intMissDirection)
        ElseIf intThrowNumber = 5 Or intThrowNumber = 6 Then
            frmGameActive.lblDice.Text = intThrowNumber & Environment.NewLine() & "So " & fucThrowCalculation(intThrowNumber)
            frmGameActive.lblMissDirection.Text = ""
            frmGameActive.lblResult.Text = fucHitCheck(intThrowNumber, intMissDirection)
        Else
            frmGameActive.lblResult.Text = "Error: Throw out of Range!"
        End If

    End Sub

    '##########################################################'
    '# Function to Return what Team player should be aming at #'
    '##########################################################'
    Function fucTeamR()
        Dim even_total As Decimal

        even_total = intPlayerRound / 2

        If even_total.ToString.Contains(".") Then
            Return "Blue"
        Else
            Return "Green"
        End If

    End Function

    '###################################################'
    '# Sub Procedure to calculate and update the socre #'
    '###################################################'
    Sub subScore(intScorePoints As Integer)
        Dim even_total As Decimal
        even_total = intPlayerRound / 2
        If even_total.ToString.Contains(".") Then
            If intScorePoints = 1 Or intScorePoints = 2 Then
                intGreenScore = intGreenScore + intScorePoints
                frmGameActive.lblGreenScore.Text = intGreenScore
            ElseIf intScorePoints = 3 Then
                frmGameActive.lblBlueScore.Text = intBlueScore
            ElseIf intScorePoints = 4 Or intScorePoints = 10 Then
                intBlueScore = intBlueScore + intScorePoints
                intGreenScore = intGreenScore - 4
                frmGameActive.lblBlueScore.Text = intBlueScore
                frmGameActive.lblGreenScore.Text = intGreenScore
            End If
        Else
            If intScorePoints = 1 Or intScorePoints = 2 Then
                intBlueScore = intBlueScore + intScorePoints
                frmGameActive.lblBlueScore.Text = intBlueScore
            ElseIf intScorePoints = 3 Then
                intGreenScore = intGreenScore + intScorePoints
                frmGameActive.lblGreenScore.Text = intGreenScore
            ElseIf intScorePoints = 4 Or intScorePoints = 10 Then
                intGreenScore = intGreenScore + intScorePoints
                intBlueScore = intBlueScore - 4
                frmGameActive.lblBlueScore.Text = intBlueScore
                frmGameActive.lblGreenScore.Text = intGreenScore
            End If
        End If
    End Sub

    '###############################################'
    '# Function to Return if a Player has been hit #'
    '###############################################'
    Function fucHitCheck(intThrowNumber As Integer, intMissDirection As Integer)
        Select Case (intThrowNumber)
            Case 1, 2, 3, 4
                strResult = fucTeamR() & fucMissHit(intThrowNumber, intMissDirection, strTarget)

                If strResult = "BlueOut of Bounds" Or strResult = "GreenOut of Bounds" Then
                    Return "Out of bounds!"
                End If
                frmGameActive.Controls.Find("txt" & strResult, True).FirstOrDefault().BackColor = Color.Red
                If frmGameActive.Controls.Find("txt" & strResult, True).FirstOrDefault().Text = "" Then
                    Return "No player found at that square!"
                ElseIf frmGameActive.Controls.Find("txt" & strResult, True).FirstOrDefault().Text <> "" Then
                    frmGameActive.btnCatch.Show()
                    frmGameActive.btnDodge.Show()
                    subScore(1)
                    Return "Possible hit on: " & frmGameActive.Controls.Find("txt" & strResult, True).FirstOrDefault().Text & ". They can now chose to Dodge or Catch!"
                Else
                    Return "Error: funHitCheck - Invalid Data"
                End If
            Case 5, 6
                If frmGameActive.Controls.Find("txt" & strTarget, True).FirstOrDefault().Text = "" Then
                    Return "No player found at that square!"
                ElseIf frmGameActive.Controls.Find("txt" & strTarget, True).FirstOrDefault().Text <> "" Then
                    frmGameActive.btnCatch.Show()
                    frmGameActive.btnDodge.Show()
                    subScore(2)
                    Return "Possible hit on: " & frmGameActive.Controls.Find("txt" & strTarget, True).FirstOrDefault().Text & ". They can now chose to Doge or Catch!"
                Else
                    Return "Error: funHitCheck - Invalid Data"
                End If
            Case Else
                Return "Error: funHitCheck - Invalid Data"
        End Select
    End Function

    '#########################################'
    '# Sub Procedure for Dodge or Catch Ball #'
    '#########################################'
    Sub subDodgeCatch(choice As String)
        If choice = "Dodge" Then
            Dim intDodgeNumber As Integer
            intDodgeNumber = random.Next(1, 7)
            subDieImage(intDodgeNumber)
            frmGameActive.lblDice.Text = intDodgeNumber
            Select Case (intThrowNumber)
                Case 1, 2

                    If intDodgeNumber <= 2 Then
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Failed to dodge the ball!"
                        frmGameActive.Controls.Find("txt" & strResult, True).FirstOrDefault().BackColor = Color.Red

                    ElseIf intDodgeNumber > 2 Then
                        subScore(3)
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Dodged the ball!"

                    End If
                Case 3, 4
                    If intDodgeNumber <= 3 Then
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Failed to dodge the ball!"
                        frmGameActive.Controls.Find("txt" & strResult, True).FirstOrDefault().BackColor = Color.Red
                    ElseIf intDodgeNumber > 3 Then
                        subScore(3)
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Doged the ball!"
                    End If
                Case 5, 6
                    If intDodgeNumber <= 4 Then
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Failed to dodge the ball!"
                        frmGameActive.Controls.Find("txt" & strTarget, True).FirstOrDefault().BackColor = Color.Red
                    ElseIf intDodgeNumber > 4 Then
                        subScore(3)
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Doged the ball!"
                    End If
            End Select
        ElseIf choice = "Catch" Then
            Dim intDodgeNumber, intCatchNumber, intCatchTotalNumber As Integer
            intDodgeNumber = random.Next(1, 7)
            intCatchNumber = random.Next(1, 9)
            intCatchTotalNumber = intDodgeNumber + intCatchNumber
            subDieImage(intDodgeNumber)
            subMissArrow(intCatchNumber)
            frmGameActive.lblDice.Text = intDodgeNumber
            frmGameActive.lblMissDirection.Text = intCatchNumber
            Select Case (intThrowNumber)
                Case 1, 2

                    If intCatchTotalNumber <= 10 Then
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Failed to catch the ball!"
                        frmGameActive.Controls.Find("txt" & strResult, True).FirstOrDefault().BackColor = Color.Red
                    ElseIf intCatchTotalNumber > 10 Then
                        subScore(5)
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Caught the ball!"

                    End If
                Case 3, 4
                    If intCatchTotalNumber <= 11 Then
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Failed to catch the ball!"
                        frmGameActive.Controls.Find("txt" & strResult, True).FirstOrDefault().BackColor = Color.Red
                    ElseIf intCatchTotalNumber > 11 Then
                        subScore(5)
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Catch the ball!"
                    End If
                Case 5, 6
                    If intCatchTotalNumber <= 12 Then
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Failed to catch the ball!"
                        frmGameActive.Controls.Find("txt" & strTarget, True).FirstOrDefault().BackColor = Color.Red
                    ElseIf intCatchTotalNumber > 12 And intCatchTotalNumber <> 14 Then
                        subScore(5)
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Catch the ball!"
                    ElseIf intCatchTotalNumber = 14 Then
                        subScore(10)
                        frmGameActive.lblResult.Text = frmGameActive.lblResult.Text & Environment.NewLine() & "Catch the ball + 10 Points!"
                    End If
            End Select

        End If
    End Sub


    '###################################################################'
    '# Functions for String Representation of Throw and Miss Direction #'
    '###################################################################'
    Function fucThrowCalculation(intThrowNumber As Integer)
        Select Case (intThrowNumber)
            Case 1, 2
                Return "2"
            Case 3, 4
                Return "1"
            Case 5, 6
                Return "Possible Hit"
            Case Else
                Return "Error"
        End Select
    End Function

    Function fucMissCalculation(intMissDirection As Integer)
        Select Case (intMissDirection)
            Case 1
                Return "North West"
            Case 2
                Return "North"
            Case 3
                Return "North East"
            Case 4
                Return "East"
            Case 5
                Return "South East"
            Case 6
                Return "South"
            Case 7
                Return "South West"
            Case 8
                Return "West"
            Case Else
                Return "Error: Miss Direction out of Expected Range!"
        End Select
    End Function

    '#########################################################################'
    '# Sub Procedures for Graphic Representation of Throw and Miss Direction #'
    '#########################################################################'
    Sub subDieImage(intThrowNumber As Integer)

        Select Case (intThrowNumber)
            Case 1
                frmGameActive.pibDice.Image = Image.FromStream(steDice1)
            Case 2
                frmGameActive.pibDice.Image = Image.FromStream(steDice2)
            Case 3
                frmGameActive.pibDice.Image = Image.FromStream(steDice3)
            Case 4
                frmGameActive.pibDice.Image = Image.FromStream(steDice4)
            Case 5
                frmGameActive.pibDice.Image = Image.FromStream(steDice5)
            Case 6
                frmGameActive.pibDice.Image = Image.FromStream(steDice6)
        End Select
    End Sub
    Sub subMissArrow(intMissDirection As Integer)
        Select Case (intMissDirection)
            Case 1
                frmGameActive.lblArrow.Text = "↖"
            Case 2
                frmGameActive.lblArrow.Text = "↑"
            Case 3
                frmGameActive.lblArrow.Text = "↗"
            Case 4
                frmGameActive.lblArrow.Text = "→"
            Case 5
                frmGameActive.lblArrow.Text = "↘"
            Case 6
                frmGameActive.lblArrow.Text = "↓"
            Case 7
                frmGameActive.lblArrow.Text = "↙"
            Case 8
                frmGameActive.lblArrow.Text = "←"
        End Select
    End Sub

    '#####################################################################'
    '# Function to return what the sqaure the player has hit if the miss #'
    '#####################################################################'
    Function fucMissHit(intThrowNumber As Integer, intMissDirection As Integer, target As String)
        target = target.Replace("Blue", "")
        target = target.Replace("Green", "")
        If intThrowNumber = 1 Or intThrowNumber = 2 Then
            Select Case (intMissDirection)
                Case 1
                    Select Case (target)
                        Case "A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2", "E1", "E2", "A3", "B3", "A4", "B4", "A5", "B5"
                            Return "Out of Bounds"

                        Case "C3"
                            Return "A1"
                        Case "C4"
                            Return "A2"
                        Case "C5"
                            Return "A3"


                        Case "D3"
                            Return "B1"
                        Case "D4"
                            Return "B2"
                        Case "D5"
                            Return "B3"

                        Case "E3"
                            Return "C1"
                        Case "E4"
                            Return "C2"
                        Case "E5"
                            Return "C3"

                        Case Else
                            Return "Error"
                    End Select
                Case 2
                    Select Case (target)
                        Case "A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2", "E1", "E2"
                            Return "Out of Bounds"

                        Case "A3"
                            Return "A1"
                        Case "A4"
                            Return "A2"
                        Case "A5"
                            Return "A3"

                        Case "B3"
                            Return "B1"
                        Case "B4"
                            Return "B2"
                        Case "B5"
                            Return "B3"


                        Case "C3"
                            Return "C1"
                        Case "C4"
                            Return "C2"
                        Case "C5"
                            Return "C3"

                        Case "D3"
                            Return "D1"
                        Case "D4"
                            Return "D2"
                        Case "D5"
                            Return "D3"

                        Case "E3"
                            Return "E1"
                        Case "E4"
                            Return "E2"
                        Case "E5"
                            Return "E3"

                        Case Else
                            Return "Error"
                    End Select
                Case 3
                    Select Case (target)
                        Case "A1", "A2", "B1", "B2", "C1", "C2", "D1", "D2", "D3", "D4", "D5", "E1", "E2", "E3", "E4", "E5"
                            Return "Out of Bounds"

                        Case "A3"
                            Return "C1"
                        Case "A4"
                            Return "C2"
                        Case "A5"
                            Return "C3"

                        Case "B3"
                            Return "D1"
                        Case "B4"
                            Return "D2"
                        Case "B5"
                            Return "D3"

                        Case "C3"
                            Return "E1"
                        Case "C4"
                            Return "E2"
                        Case "C5"
                            Return "E3"

                        Case Else
                            Return "Error"
                    End Select
                Case 4
                    Select Case (target)
                        Case "D1", "D2", "D3", "D4", "D5", "E1", "E2", "E3", "E4", "E5"
                            Return "Out of Bounds"

                        Case "A1"
                            Return "C1"
                        Case "A2"
                            Return "C2"
                        Case "A3"
                            Return "C3"
                        Case "A4"
                            Return "C4"
                        Case "A5"
                            Return "C5"

                        Case "B1"
                            Return "D1"
                        Case "B2"
                            Return "D2"
                        Case "B3"
                            Return "D3"
                        Case "B4"
                            Return "D4"
                        Case "B5"
                            Return "D5"

                        Case "C1"
                            Return "E1"
                        Case "C2"
                            Return "E2"
                        Case "C3"
                            Return "E3"
                        Case "C4"
                            Return "E4"
                        Case "C5"
                            Return "E5"

                        Case Else
                            Return "Error"
                    End Select
                Case 5
                    Select Case (target)
                        Case "A4", "A5", "B4", "B5", "C4", "C5", "D3", "D4", "D5", "E4", "E5", "D1", "D2", "E1", "E2", "E3"
                            Return "Out of Bounds"

                        Case "A1"
                            Return "C3"
                        Case "A2"
                            Return "C4"
                        Case "A3"
                            Return "C5"

                        Case "B1"
                            Return "D3"
                        Case "B2"
                            Return "D4"
                        Case "B3"
                            Return "D5"

                        Case "C1"
                            Return "E3"
                        Case "C2"
                            Return "E4"
                        Case "C3"
                            Return "E5"

                        Case Else
                            Return "Error"
                    End Select
                Case 6
                    Select Case (target)
                        Case "A4", "A5", "B4", "B5", "C4", "C5", "D4", "D5", "E4", "E5"
                            Return "Out of Bounds"

                        Case "A1"
                            Return "A3"
                        Case "A2"
                            Return "A4"
                        Case "A3"
                            Return "A5"

                        Case "B1"
                            Return "B3"
                        Case "B2"
                            Return "B4"
                        Case "B3"
                            Return "B5"

                        Case "C1"
                            Return "C3"
                        Case "C2"
                            Return "C4"
                        Case "C3"
                            Return "C5"

                        Case "D1"
                            Return "D3"
                        Case "D2"
                            Return "D4"
                        Case "D3"
                            Return "D5"

                        Case "E1"
                            Return "E3"
                        Case "E2"
                            Return "E4"
                        Case "E3"
                            Return "E5"

                        Case Else
                            Return "Error"
                    End Select
                Case 7
                    Select Case (target)
                        Case "A1", "B1", "A2", "B2", "A3", "B3", "A4", "B4", "A5", "B5", "C4", "C5", "D4", "D5", "E4", "E5"
                            Return "Out of Bounds"
                        Case "C1"
                            Return "A3"
                        Case "C2"
                            Return "A4"
                        Case "C3"
                            Return "A5"

                        Case "D1"
                            Return "B3"
                        Case "D2"
                            Return "B4"
                        Case "D3"
                            Return "B5"

                        Case "E1"
                            Return "C3"
                        Case "E2"
                            Return "C4"
                        Case "E3"
                            Return "C5"

                        Case Else
                            Return "Error"
                    End Select
                Case 8
                    Select Case (target)
                        Case "A1", "B1", "A2", "B2", "A3", "B3", "A4", "B4", "A5", "B5"
                            Return "Out of Bounds"

                        Case "C1"
                            Return "A1"
                        Case "C2"
                            Return "A2"
                        Case "C3"
                            Return "A3"
                        Case "C4"
                            Return "A4"
                        Case "C5"
                            Return "A5"

                        Case "D1"
                            Return "B1"
                        Case "D2"
                            Return "B2"
                        Case "D3"
                            Return "B3"
                        Case "D4"
                            Return "B4"
                        Case "D5"
                            Return "B5"

                        Case "E1"
                            Return "C1"
                        Case "E2"
                            Return "C2"
                        Case "E3"
                            Return "C3"
                        Case "E4"
                            Return "C4"
                        Case "E5"
                            Return "C5"


                        Case Else
                            Return "Error"
                    End Select
            End Select
        ElseIf intThrowNumber = 3 Or intThrowNumber = 4 Then
            Select Case (intMissDirection)
                Case 1
                    Select Case (target)
                        Case "A1", "A2", "A3", "A4", "A5", "B1", "C1", "D1", "E1"
                            Return "Out of Bounds"

                        Case "B5"
                            Return "A4"
                        Case "B4"
                            Return "A3"
                        Case "B3"
                            Return "A2"
                        Case "B2"
                            Return "A1"

                        Case "C5"
                            Return "B4"
                        Case "C4"
                            Return "B3"
                        Case "C3"
                            Return "B2"
                        Case "C2"
                            Return "B1"

                        Case "D5"
                            Return "C4"
                        Case "D4"
                            Return "C3"
                        Case "D3"
                            Return "C2"
                        Case "D2"
                            Return "C1"

                        Case "E5"
                            Return "D4"
                        Case "E4"
                            Return "D3"
                        Case "E3"
                            Return "D2"
                        Case "E2"
                            Return "D1"

                        Case Else
                            Return "Error"
                    End Select
                Case 2

                    Select Case (target)
                        Case "A1", "B1", "C1", "D1", "E1"
                            Return "Out of Bounds"

                        Case "A2"
                            Return "A1"
                        Case "A3"
                            Return "A2"
                        Case "A4"
                            Return "A3"
                        Case "A5"
                            Return "A4"

                        Case "B2"
                            Return "B1"
                        Case "B3"
                            Return "B2"
                        Case "B4"
                            Return "B3"
                        Case "B5"
                            Return "B4"

                        Case "C2"
                            Return "C1"
                        Case "C3"
                            Return "C2"
                        Case "C4"
                            Return "C3"
                        Case "C5"
                            Return "C4"

                        Case "D2"
                            Return "D1"
                        Case "D3"
                            Return "D2"
                        Case "D4"
                            Return "D3"
                        Case "D5"
                            Return "D4"

                        Case "E2"
                            Return "E1"
                        Case "E3"
                            Return "E2"
                        Case "E4"
                            Return "E3"
                        Case "E5"
                            Return "E4"

                        Case Else
                            Return "Error"
                    End Select
                Case 3
                    Select Case (target)
                        Case "E1", "E2", "E3", "E4", "E5", "A1", "B1", "C1", "D1"
                            Return "Out of Bounds"

                        Case "A2"
                            Return "B1"
                        Case "A3"
                            Return "B2"
                        Case "A4"
                            Return "B3"
                        Case "A5"
                            Return "B4"

                        Case "B2"
                            Return "C1"
                        Case "B3"
                            Return "C2"
                        Case "B4"
                            Return "C3"
                        Case "B5"
                            Return "C4"

                        Case "C2"
                            Return "D1"
                        Case "C3"
                            Return "D2"
                        Case "C4"
                            Return "D3"
                        Case "C5"
                            Return "D4"

                        Case "D2"
                            Return "E1"
                        Case "D3"
                            Return "E2"
                        Case "D4"
                            Return "E3"
                        Case "D5"
                            Return "E4"

                        Case Else
                            Return "Error"
                    End Select
                Case 4
                    Select Case (target)
                        Case "E1", "E2", "E3", "E4", "E5"
                            Return "Out of Bounds"

                        Case "A1"
                            Return "B1"
                        Case "A2"
                            Return "B2"
                        Case "A3"
                            Return "B3"
                        Case "A4"
                            Return "B4"
                        Case "A5"
                            Return "B5"

                        Case "B1"
                            Return "C1"
                        Case "B2"
                            Return "C2"
                        Case "B3"
                            Return "C3"
                        Case "B4"
                            Return "C4"
                        Case "B5"
                            Return "C5"

                        Case "C1"
                            Return "D1"
                        Case "C2"
                            Return "D2"
                        Case "C3"
                            Return "D3"
                        Case "C4"
                            Return "D4"
                        Case "C5"
                            Return "D5"

                        Case "D1"
                            Return "E1"
                        Case "D2"
                            Return "E2"
                        Case "D3"
                            Return "E3"
                        Case "D4"
                            Return "E4"
                        Case "D5"
                            Return "E5"

                        Case Else
                            Return "Error"
                    End Select
                Case 5
                    Select Case (target)
                        Case "E1", "E2", "E3", "E4", "E5", "A5", "B5", "C5", "D5", "E5"
                            Return "Out of Bounds"

                        Case "A1"
                            Return "B2"
                        Case "A2"
                            Return "B3"
                        Case "A3"
                            Return "B4"
                        Case "A4"
                            Return "B5"

                        Case "B1"
                            Return "C2"
                        Case "B2"
                            Return "C3"
                        Case "B3"
                            Return "C4"
                        Case "B4"
                            Return "C5"

                        Case "C1"
                            Return "D2"
                        Case "C2"
                            Return "D3"
                        Case "C3"
                            Return "D4"
                        Case "C4"
                            Return "D5"

                        Case "D1"
                            Return "E2"
                        Case "D2"
                            Return "E3"
                        Case "D3"
                            Return "E4"
                        Case "D4"
                            Return "E5"

                        Case Else
                            Return "Error"
                    End Select
                Case 6
                    Select Case (target)
                        Case "A5", "B5", "C5", "D5", "E5"
                            Return "Out of Bounds"

                        Case "A1"
                            Return "A2"
                        Case "A2"
                            Return "A3"
                        Case "A3"
                            Return "A4"
                        Case "A4"
                            Return "A5"

                        Case "B1"
                            Return "B2"
                        Case "B2"
                            Return "B3"
                        Case "B3"
                            Return "B4"
                        Case "B4"
                            Return "B5"

                        Case "C1"
                            Return "C2"
                        Case "C2"
                            Return "C3"
                        Case "C3"
                            Return "C4"
                        Case "C4"
                            Return "C5"

                        Case "D1"
                            Return "D2"
                        Case "D2"
                            Return "D3"
                        Case "D3"
                            Return "D4"
                        Case "D4"
                            Return "D5"

                        Case "E1"
                            Return "E2"
                        Case "E2"
                            Return "E3"
                        Case "E3"
                            Return "E4"
                        Case "E4"
                            Return "E5"

                        Case Else
                            Return "Error"
                    End Select
                Case 7
                    Select Case (target)
                        Case "A1", "A2", "A3", "A4", "A5", "B5", "C5", "D5", "E5"
                            Return "Out of Bounds"

                        Case "B1"
                            Return "A2"
                        Case "B2"
                            Return "A3"
                        Case "B3"
                            Return "A4"
                        Case "B4"
                            Return "A5"

                        Case "C1"
                            Return "B2"
                        Case "C2"
                            Return "B3"
                        Case "C3"
                            Return "B4"
                        Case "C4"
                            Return "B5"

                        Case "D1"
                            Return "C2"
                        Case "D2"
                            Return "C3"
                        Case "D3"
                            Return "C4"
                        Case "D4"
                            Return "C5"

                        Case "E1"
                            Return "D2"
                        Case "E2"
                            Return "D3"
                        Case "E3"
                            Return "D4"
                        Case "E4"
                            Return "D5"

                        Case Else
                            Return "Error"
                    End Select
                Case 8
                    Select Case (target)
                        Case "A1", "A2", "A3", "A4", "A5"
                            Return "Out of Bounds"

                        Case "B1"
                            Return "A1"
                        Case "B2"
                            Return "A2"
                        Case "B3"
                            Return "A3"
                        Case "B4"
                            Return "A4"
                        Case "B5"
                            Return "A5"

                        Case "C1"
                            Return "B1"
                        Case "C2"
                            Return "B2"
                        Case "C3"
                            Return "B3"
                        Case "C4"
                            Return "B4"
                        Case "C5"
                            Return "B5"

                        Case "D1"
                            Return "C1"
                        Case "D2"
                            Return "C2"
                        Case "D3"
                            Return "C3"
                        Case "D4"
                            Return "C4"
                        Case "D5"
                            Return "C5"

                        Case "E1"
                            Return "D1"
                        Case "E2"
                            Return "D2"
                        Case "E3"
                            Return "D3"
                        Case "E4"
                            Return "D4"
                        Case "E5"
                            Return "D5"

                        Case Else
                            Return "Error"
                    End Select
            End Select
        Else
            Return "Error"
        End If
        Return "Error"
    End Function

    '###################################################'
    '# Sub Procedure to Reset all textboxes and Labels #'
    '###################################################'
    Sub subReset()
        frmGameActive.txtBlueA1.BackColor = Color.White
        frmGameActive.txtBlueA2.BackColor = Color.White
        frmGameActive.txtBlueA3.BackColor = Color.White
        frmGameActive.txtBlueA4.BackColor = Color.White
        frmGameActive.txtBlueA5.BackColor = Color.White
        frmGameActive.txtBlueB1.BackColor = Color.White
        frmGameActive.txtBlueB2.BackColor = Color.White
        frmGameActive.txtBlueB3.BackColor = Color.White
        frmGameActive.txtBlueB4.BackColor = Color.White
        frmGameActive.txtBlueB5.BackColor = Color.White
        frmGameActive.txtBlueC1.BackColor = Color.White
        frmGameActive.txtBlueC2.BackColor = Color.White
        frmGameActive.txtBlueC3.BackColor = Color.White
        frmGameActive.txtBlueC4.BackColor = Color.White
        frmGameActive.txtBlueC5.BackColor = Color.White
        frmGameActive.txtBlueD1.BackColor = Color.White
        frmGameActive.txtBlueD2.BackColor = Color.White
        frmGameActive.txtBlueD3.BackColor = Color.White
        frmGameActive.txtBlueD4.BackColor = Color.White
        frmGameActive.txtBlueD5.BackColor = Color.White
        frmGameActive.txtBlueE1.BackColor = Color.White
        frmGameActive.txtBlueE2.BackColor = Color.White
        frmGameActive.txtBlueE3.BackColor = Color.White
        frmGameActive.txtBlueE4.BackColor = Color.White
        frmGameActive.txtBlueE5.BackColor = Color.White

        frmGameActive.txtGreenA1.BackColor = Color.White
        frmGameActive.txtGreenA2.BackColor = Color.White
        frmGameActive.txtGreenA3.BackColor = Color.White
        frmGameActive.txtGreenA4.BackColor = Color.White
        frmGameActive.txtGreenA5.BackColor = Color.White
        frmGameActive.txtGreenB1.BackColor = Color.White
        frmGameActive.txtGreenB2.BackColor = Color.White
        frmGameActive.txtGreenB3.BackColor = Color.White
        frmGameActive.txtGreenB4.BackColor = Color.White
        frmGameActive.txtGreenB5.BackColor = Color.White
        frmGameActive.txtGreenC1.BackColor = Color.White
        frmGameActive.txtGreenC2.BackColor = Color.White
        frmGameActive.txtGreenC3.BackColor = Color.White
        frmGameActive.txtGreenC4.BackColor = Color.White
        frmGameActive.txtGreenC5.BackColor = Color.White
        frmGameActive.txtGreenD1.BackColor = Color.White
        frmGameActive.txtGreenD2.BackColor = Color.White
        frmGameActive.txtGreenD3.BackColor = Color.White
        frmGameActive.txtGreenD4.BackColor = Color.White
        frmGameActive.txtGreenD5.BackColor = Color.White
        frmGameActive.txtGreenE1.BackColor = Color.White
        frmGameActive.txtGreenE2.BackColor = Color.White
        frmGameActive.txtGreenE3.BackColor = Color.White
        frmGameActive.txtGreenE4.BackColor = Color.White
        frmGameActive.txtGreenE5.BackColor = Color.White

        frmGameActive.lblPlayer.Text = ""
        frmGameActive.lblTargetGrid.Text = ""
        frmGameActive.lblTargetPlayer.Text = ""
        frmGameActive.lblDice.Text = ""
        frmGameActive.lblMissDirection.Text = ""
        frmGameActive.lblResult.Text = ""


        strTarget = ""
        intThrowNumber = 0
        intMissDirection = 0
        strResult = ""

        frmGameActive.btnDodge.Hide()
        frmGameActive.btnCatch.Hide()

        frmGameActive.btnTarget.Show()
        frmGameActive.btnPlayer.Hide()
    End Sub

    Sub subExit()
        Select Case MsgBox("Are you sure you want to Exit?", vbYesNo, "Dodgeball Online - Exit Confirmation")
            Case MsgBoxResult.Yes
                Application.Exit()
            Case MsgBoxResult.No
                Exit Sub
        End Select
    End Sub

    Sub subHelp()
        Dim webAddress As String = "https://dodgeball.bxdavies.co.uk/docs/"
        Process.Start(webAddress)
    End Sub

    Sub subInfo()
        Dim mymessage As String = "Dodgeball Online" & Environment.NewLine & "Version: 1.1" & Environment.NewLine & "Github: www.github.com/dodgeball-online/dodgeball-online" & Environment.NewLine & "This Program and its Source code is Licensed under General Public License v3.0, for more information see GitHub repository." & Environment.NewLine & Environment.NewLine & "Attribution: " & Environment.NewLine() & "Dice Images: Made by Alfredo Hernandez from www.flaticon.com" & Environment.NewLine & "Icons: Made by Google from www.flaticon.com"
        MsgBox(mymessage)
    End Sub
End Module