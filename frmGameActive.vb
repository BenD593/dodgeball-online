Imports System.Resources
Imports System.IO
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Public Class frmGameActive
    Public lisBlueTeam As New List(Of String) From {strBlueA1, strBlueA2, strBlueA3, strBlueA4, strBlueA5, strBlueB1, strBlueB2, strBlueB3, strBlueB4, strBlueB5, strBlueC1, strBlueC2, strBlueC3, strBlueC4, strBlueC5, strBlueD1, strBlueD2, strBlueD3, strBlueD4, strBlueD5, strBlueE1, strBlueE2, strBlueE3, strBlueE4, strBlueE5}
    Public lisGreenTeam As New List(Of String) From {strGreenA1, strGreenA2, strGreenA3, strGreenA4, strGreenA5, strGreenB1, strGreenB2, strGreenB3, strGreenB4, strGreenB5, strGreenC1, strGreenC2, strGreenC3, strGreenC4, strGreenC5, strGreenD1, strGreenD2, strGreenD3, strGreenD4, strGreenD5, strGreenE1, strGreenE2, strGreenE3, strGreenE4, strGreenE5}
    Public lisAllPlayers As New List(Of String)
    Public lisDiceImages As New List(Of Stream) From {steDice1, steDice2, steDice2, steDice3, steDice4, steDice5, steDice6}
    Public lisArrows As New List(Of String) From {"←", "↑", "→", "↓", "↖", "↗", "↘", "↙"}
    Public intDieCount, intArrowCount As Integer

    Private Sub frmGameActive_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '##################'
        '# Visual Changes #'
        '##################'
        ' Maximise Window '
        WindowState = 2

        ' Hide Buttons '
        btnPlayer.Hide()
        btnThrow.Hide()
        btnTarget.Hide()
        btnDodge.Hide()
        btnCatch.Hide()


        '###########################################'
        '# Fill values into TextBox from Variables #'
        '###########################################'
        ' Blue Team'
        txtBlueA1.Text = strBlueA1
        txtBlueA2.Text = strBlueA2
        txtBlueA3.Text = strBlueA3
        txtBlueA4.Text = strBlueA4
        txtBlueA5.Text = strBlueA5
        txtBlueB1.Text = strBlueB1
        txtBlueB2.Text = strBlueB2
        txtBlueB3.Text = strBlueB3
        txtBlueB4.Text = strBlueB4
        txtBlueB5.Text = strBlueB5
        txtBlueC1.Text = strBlueC1
        txtBlueC2.Text = strBlueC2
        txtBlueC3.Text = strBlueC3
        txtBlueC4.Text = strBlueC4
        txtBlueC5.Text = strBlueC5
        txtBlueD1.Text = strBlueD1
        txtBlueD2.Text = strBlueD2
        txtBlueD3.Text = strBlueD3
        txtBlueD4.Text = strBlueD4
        txtBlueD5.Text = strBlueD5
        txtBlueE1.Text = strBlueE1
        txtBlueE2.Text = strBlueE2
        txtBlueE3.Text = strBlueE3
        txtBlueE4.Text = strBlueE4
        txtBlueE5.Text = strBlueE5

        ' Green Team '
        txtGreenA1.Text = strGreenA1
        txtGreenA2.Text = strGreenA2
        txtGreenA3.Text = strGreenA3
        txtGreenA4.Text = strGreenA4
        txtGreenA5.Text = strGreenA5
        txtGreenB1.Text = strGreenB1
        txtGreenB2.Text = strGreenB2
        txtGreenB3.Text = strGreenB3
        txtGreenB4.Text = strGreenB4
        txtGreenB5.Text = strGreenB5
        txtGreenC1.Text = strGreenC1
        txtGreenC2.Text = strGreenC2
        txtGreenC3.Text = strGreenC3
        txtGreenC4.Text = strGreenC4
        txtGreenC5.Text = strGreenC5
        txtGreenD1.Text = strGreenD1
        txtGreenD2.Text = strGreenD2
        txtGreenD3.Text = strGreenD3
        txtGreenD4.Text = strGreenD4
        txtGreenD5.Text = strGreenD5
        txtGreenE1.Text = strGreenE1
        txtGreenE2.Text = strGreenE2
        txtGreenE3.Text = strGreenE3
        txtGreenE4.Text = strGreenE4
        txtGreenE5.Text = strGreenE5

        '####################################'
        '# Remove empty Elements from Lists #'
        '####################################'
        lisBlueTeam.RemoveAll(Function(str) String.IsNullOrWhiteSpace(str))
        lisGreenTeam.RemoveAll(Function(str) String.IsNullOrWhiteSpace(str))

        '##########################'
        '# Create List AllPlayers #'
        '##########################'
        Dim intAllPlayerCounter As Integer = 0
        Do
            lisAllPlayers.Add(lisBlueTeam(intAllPlayerCounter))
            lisAllPlayers.Add(lisGreenTeam(intAllPlayerCounter))
            intAllPlayerCounter += 1
        Loop Until intAllPlayerCounter >= lisBlueTeam.Count Or intAllPlayerCounter >= lisGreenTeam.Count()

        '###########################################################################'
        '# Loop through both Lists adding each item to correct Teams Team Text Box #'
        '###########################################################################'
        Dim intListCount As Integer = 0
        Dim intListItem As Integer = 1
        Do
            rtxtBlueTeam.AppendText(Environment.NewLine() & intListItem & ". " & lisBlueTeam(intListCount))
            intListItem = intListItem + 1
            rtxtGreenTeam.AppendText(Environment.NewLine() & intListItem & ". " & lisGreenTeam(intListCount))
            intListItem = intListItem + 1
            intListCount = intListCount + 1
        Loop Until intListCount + 1 > lisAllPlayers.Count / 2

    End Sub

    Private Sub btnRound_Click(sender As Object, e As EventArgs) Handles btnRound.Click
        intRound += 1
        intPlayerRound = -1
        btnRound.Text = "Next Round"
        lblRound.Text = intRound
        btnRound.Hide()
        btnPlayer.Show()

    End Sub

    Private Sub btnPlayer_Click(sender As Object, e As EventArgs) Handles btnPlayer.Click
        btnPlayer.Text = "Next Player"
        '##########################################################################'
        '# Check if All Players have had a go. Else update Labels for Next Player #'
        '##########################################################################'
        If intPlayerRound + 1 >= lisAllPlayers.Count Then
            btnPlayer.Text = "First Player"
            subReset()

            btnRound.Show()
            btnPlayer.Hide()
            btnThrow.Hide()
            btnTarget.Hide()


        ElseIf intPlayerRound <= lisAllPlayers.Count And intPlayerRound = -1 Then
            subReset()

            intPlayerRound += 1
            lblPlayerRound.Text = intPlayerRound + 1 & " / " & lisAllPlayers.Count
            lblPlayer.Text = lisAllPlayers(intPlayerRound)
        ElseIf intPlayerRound <= lisAllPlayers.Count Then
            subReset()

            intPlayerRound += 1
            lblPlayerRound.Text = intPlayerRound + 1 & " / " & lisAllPlayers.Count
            lblPlayer.Text = lisAllPlayers(intPlayerRound)
        End If
    End Sub

    Private Sub btnTarget_Click(sender As Object, e As EventArgs) Handles btnTarget.Click
        '#################################'
        '# Check if a target is Selected #'
        '#################################'
        If strTarget = "" Then
            MsgBox("No target selected!" & Environment.NewLine() & "Please select a target and try again!", vbExclamation, "Dodegball on Zoom - Error")
            Exit Sub
        End If

        '###########################################'
        '# Check if target is on the opposite team #'
        '###########################################'
        If strTarget.Contains(fucTeamR) = False Then
            MsgBox("Target on wrong grid!" & Environment.NewLine() & "Please select a target from the " & fucTeamR() & " grid and try again!", vbExclamation, "Dodegball on Zoom - Error")
            Exit Sub
        End If

        lblTargetGrid.Text = strTarget

        '########################################'
        '# Check if there is a player at target #'
        '########################################'
        Dim txtCellControl = Controls.Find("txt" & strTarget, True).FirstOrDefault()
        If txtCellControl.Text = "" Then
            lblTargetPlayer.Text = "No player found!"
        Else
            lblTargetPlayer.Text = txtCellControl.Text
        End If

        Controls.Find("txt" & strTarget, True).FirstOrDefault().BackColor = Color.Orange
        btnThrow.Show()
    End Sub

    Private Sub btnThrow_Click(sender As Object, e As EventArgs) Handles btnThrow.Click
        intDieCount = 0
        intArrowCount = 0
        timThrow.Enabled = True
        btnTarget.Hide()
        btnThrow.Hide()

    End Sub

    Private Sub btnDodge_Click(sender As Object, e As EventArgs) Handles btnDodge.Click
        intDieCount = 0
        intArrowCount = 0
        timDodge.Enabled = True
    End Sub

    Private Sub btnCatch_Click(sender As Object, e As EventArgs) Handles btnCatch.Click
        intDieCount = 0
        intArrowCount = 0
        timCatch.Enabled = True
    End Sub

    '########################################################'
    '# Timers to handle dice roll images and arrow rotation #'
    '########################################################'
    Private Sub timThrow_Tick(sender As Object, e As EventArgs) Handles timThrow.Tick
        pibDice.Image = Image.FromStream(lisDiceImages(random.Next(0, 6)))
        intDieCount = intDieCount + 1
        lblArrow.Text = lisArrows(random.Next(0, 8))
        intArrowCount = intArrowCount + 1
        If intArrowCount > 8 And intDieCount > 7 Then
            intArrowCount = 0
            timThrow.Enabled = False
            btnPlayer.Show()
            subResult()
        End If
    End Sub

    Private Sub timDodge_Tick(sender As Object, e As EventArgs) Handles timDodge.Tick
        pibDice.Image = Image.FromStream(lisDiceImages(random.Next(0, 6)))
        intDieCount = intDieCount + 1
        If intDieCount > 7 Then
            intDieCount = 0
            timDodge.Enabled = False
            btnPlayer.Show()
            subDodgeCatch("Dodge")
        End If
    End Sub

    Private Sub timCatch_Tick(sender As Object, e As EventArgs) Handles timCatch.Tick
        pibDice.Image = Image.FromStream(lisDiceImages(random.Next(0, 6)))
        intDieCount = intDieCount + 1
        lblArrow.Text = lisArrows(random.Next(0, 8))
        intArrowCount = intArrowCount + 1
        If intArrowCount > 8 And intDieCount > 7 Then
            intArrowCount = 0
            timCatch.Enabled = False
            btnPlayer.Show()
            subDodgeCatch("Catch")
        End If
    End Sub

    '#######################################################'
    '# Event Handlers of when a text box is double clicked #'
    '#######################################################'

    ' Blue Team '
    Private Sub txtBlueA1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueA1.DoubleClick
        strTarget = "BlueA1"
    End Sub

    Private Sub txtBlueA2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueA2.DoubleClick
        strTarget = "BlueA2"
    End Sub

    Private Sub txtBlueA3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueA3.DoubleClick
        strTarget = "BlueA3"
    End Sub
    Private Sub txtBlueA4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueA4.DoubleClick
        strTarget = "BlueA4"
    End Sub
    Private Sub txtBlueA5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueA5.DoubleClick
        strTarget = "BlueA5"
    End Sub
    Private Sub txtBlueB1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueB1.DoubleClick
        strTarget = "BlueB1"
    End Sub
    Private Sub txtBlueB2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueB2.DoubleClick
        strTarget = "BlueB2"
    End Sub
    Private Sub txtBlueB3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueB3.DoubleClick
        strTarget = "BlueB3"
    End Sub
    Private Sub txtBlueB4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueB4.DoubleClick
        strTarget = "BlueB4"
    End Sub
    Private Sub txtBlueB5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueB5.DoubleClick
        strTarget = "BlueB5"
    End Sub
    Private Sub txtBlueC1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueC1.DoubleClick
        strTarget = "BlueC1"
    End Sub
    Private Sub txtBlueC2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueC2.DoubleClick
        strTarget = "BlueC2"
    End Sub
    Private Sub txtBlueC3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueC3.DoubleClick
        strTarget = "BlueC3"
    End Sub
    Private Sub txtBlueC4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueC4.DoubleClick
        strTarget = "BlueC4"
    End Sub
    Private Sub txtBlueC5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueC5.DoubleClick
        strTarget = "BlueC5"
    End Sub
    Private Sub txtBlueD1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueD1.DoubleClick
        strTarget = "BlueD1"
    End Sub
    Private Sub txtBlueD2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueD2.DoubleClick
        strTarget = "BlueD2"
    End Sub
    Private Sub txtBlueD3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueD3.DoubleClick
        strTarget = "BlueD3"
    End Sub
    Private Sub txtBlueD4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueD4.DoubleClick
        strTarget = "BlueD4"
    End Sub
    Private Sub txtBlueD5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueD5.DoubleClick
        strTarget = "BlueD5"
    End Sub
    Private Sub txtBlueE1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueE1.DoubleClick
        strTarget = "BlueE1"
    End Sub
    Private Sub txtBlueE2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueE2.DoubleClick
        strTarget = "BlueE2"
    End Sub
    Private Sub txtBlueE3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueE3.DoubleClick
        strTarget = "BlueE3"
    End Sub
    Private Sub txtBlueE4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueE4.DoubleClick
        strTarget = "BlueE4"
    End Sub
    Private Sub txtBlueE5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtBlueE5.DoubleClick
        strTarget = "BlueE5"
    End Sub

    ' Green Team '
    Private Sub txtGreenA1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenA1.DoubleClick
        strTarget = "GreenA1"
    End Sub
    Private Sub txtGreenA2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenA2.DoubleClick
        strTarget = "GreenA2"
    End Sub
    Private Sub txtGreenA3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenA3.DoubleClick
        strTarget = "GreenA3"
    End Sub
    Private Sub txtGreenA4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenA4.DoubleClick
        strTarget = "GreenA4"
    End Sub
    Private Sub txtGreenA5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenA5.DoubleClick
        strTarget = "GreenA5"
    End Sub
    Private Sub txtGreenB1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenB1.DoubleClick
        strTarget = "GreenB1"
    End Sub
    Private Sub txtGreenB2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenB2.DoubleClick
        strTarget = "GreenB2"
    End Sub
    Private Sub txtGreenB3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenB3.DoubleClick
        strTarget = "GreenB3"
    End Sub
    Private Sub txtGreenB4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenB4.DoubleClick
        strTarget = "GreenB4"
    End Sub
    Private Sub txtGreenB5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenB5.DoubleClick
        strTarget = "GreenB5"
    End Sub
    Private Sub txtGreenC1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenC1.DoubleClick
        strTarget = "GreenC1"
    End Sub
    Private Sub txtGreenC2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenC2.DoubleClick
        strTarget = "GreenC2"
    End Sub
    Private Sub txtGreenC3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenC3.DoubleClick
        strTarget = "GreenC3"
    End Sub
    Private Sub txtGreenC4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenC4.DoubleClick
        strTarget = "GreenC4"
    End Sub
    Private Sub txtGreenC5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenC5.DoubleClick
        strTarget = "GreenC5"
    End Sub
    Private Sub txtGreenD1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenD1.DoubleClick
        strTarget = "GreenD1"
    End Sub
    Private Sub txtGreenD2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenD2.DoubleClick
        strTarget = "GreenD2"
    End Sub
    Private Sub txtGreenD3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenD3.DoubleClick
        strTarget = "GreenD3"
    End Sub
    Private Sub txtGreenD4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenD4.DoubleClick
        strTarget = "GreenD4"
    End Sub
    Private Sub txtGreenD5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenD5.DoubleClick
        strTarget = "GreenD5"
    End Sub
    Private Sub txtGreenE1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenE1.DoubleClick
        strTarget = "GreenE1"
    End Sub
    Private Sub txtGreenE2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenE2.DoubleClick
        strTarget = "GreenE2"
    End Sub
    Private Sub txtGreenE3_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenE3.DoubleClick
        strTarget = "GreenE3"
    End Sub
    Private Sub txtGreenE4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenE4.DoubleClick
        strTarget = "GreenE4"
    End Sub
    Private Sub txtGreenE5_Click(ByVal sender As Object, ByVal e As EventArgs) Handles txtGreenE5.DoubleClick
        strTarget = "GreenE5"
    End Sub

    Private Sub tsbExit_Click(sender As Object, e As EventArgs) Handles tsbExit.Click
        subExit()
    End Sub
    Private Sub tsbHelp_Click(sender As Object, e As EventArgs) Handles tsbHelp.Click
        subHelp()
    End Sub
    Private Sub tsbInfo_Click(sender As Object, e As EventArgs) Handles tsbInfo.Click
        subInfo()
    End Sub

End Class