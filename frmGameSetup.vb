Imports System.IO
Imports System.Reflection
Public Class frmGameSetup
    Private Sub frmGameSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = 2
    End Sub
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        strBlueA1 = txtBlueA1.Text
        strBlueA2 = txtBlueA2.Text
        strBlueA3 = txtBlueA3.Text
        strBlueA4 = txtBlueA4.Text
        strBlueA5 = txtBlueA5.Text
        strBlueB1 = txtBlueB1.Text
        strBlueB2 = txtBlueB2.Text
        strBlueB3 = txtBlueB3.Text
        strBlueB4 = txtBlueB4.Text
        strBlueB5 = txtBlueB5.Text
        strBlueC1 = txtBlueC1.Text
        strBlueC2 = txtBlueC2.Text
        strBlueC3 = txtBlueC3.Text
        strBlueC4 = txtBlueC4.Text
        strBlueC5 = txtBlueC5.Text
        strBlueD1 = txtBlueD1.Text
        strBlueD2 = txtBlueD2.Text
        strBlueD3 = txtBlueD3.Text
        strBlueD4 = txtBlueD4.Text
        strBlueD5 = txtBlueD5.Text
        strBlueE1 = txtBlueE1.Text
        strBlueE2 = txtBlueE2.Text
        strBlueE3 = txtBlueE3.Text
        strBlueE4 = txtBlueE4.Text
        strBlueE5 = txtBlueE5.Text

        strGreenA1 = txtGreenA1.Text
        strGreenA2 = txtGreenA2.Text
        strGreenA3 = txtGreenA3.Text
        strGreenA4 = txtGreenA4.Text
        strGreenA5 = txtGreenA5.Text
        strGreenB1 = txtGreenB1.Text
        strGreenB2 = txtGreenB2.Text
        strGreenB3 = txtGreenB3.Text
        strGreenB4 = txtGreenB4.Text
        strGreenB5 = txtGreenB5.Text
        strGreenC1 = txtGreenC1.Text
        strGreenC2 = txtGreenC2.Text
        strGreenC3 = txtGreenC3.Text
        strGreenC4 = txtGreenC4.Text
        strGreenC5 = txtGreenC5.Text
        strGreenD1 = txtGreenD1.Text
        strGreenD2 = txtGreenD2.Text
        strGreenD3 = txtGreenD3.Text
        strGreenD4 = txtGreenD4.Text
        strGreenD5 = txtGreenD5.Text
        strGreenE1 = txtGreenE1.Text
        strGreenE2 = txtGreenE2.Text
        strGreenE3 = txtGreenE3.Text
        strGreenE4 = txtGreenE4.Text
        strGreenE5 = txtGreenE5.Text


        Dim lisBlueTeam As New List(Of String) From {strBlueA1, strBlueA2, strBlueA3, strBlueA4, strBlueA5, strBlueB1, strBlueB2, strBlueB3, strBlueB4, strBlueB5, strBlueC1, strBlueC2, strBlueC3, strBlueC4, strBlueC5, strBlueD1, strBlueD2, strBlueD3, strBlueD4, strBlueD5, strBlueE1, strBlueE2, strBlueE3, strBlueE4, strBlueE5}
        Dim lisGreenTeam As New List(Of String) From {strGreenA1, strGreenA2, strGreenA3, strGreenA4, strGreenA5, strGreenB1, strGreenB2, strGreenB3, strGreenB4, strGreenB5, strGreenC1, strGreenC2, strGreenC3, strGreenC4, strGreenC5, strGreenD1, strGreenD2, strGreenD3, strGreenD4, strGreenD5, strGreenE1, strGreenE2, strGreenE3, strGreenE4, strGreenE5}

        lisBlueTeam.RemoveAll(Function(str) String.IsNullOrWhiteSpace(str))
        lisGreenTeam.RemoveAll(Function(str) String.IsNullOrWhiteSpace(str))

        If lisGreenTeam.Count = 0 Or lisBlueTeam.Count = 0 Then
            MsgBox("No players in either the Blue team or Green team! " & Environment.NewLine() & "Please add players and try again!", vbExclamation, "Dodegball Online - Error")
            Exit Sub
        End If


        If lisGreenTeam.Count <> lisBlueTeam.Count Then
            MsgBox("Teams are not equal! " & Environment.NewLine() & "Please add players and try again!", vbExclamation, "Dodegball Online - Error")
            Exit Sub
        End If

        frmGameActive.Activate()
        frmGameActive.Show()
        Me.Hide()

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
