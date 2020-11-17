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

        frmGameActive.Activate()
        frmGameActive.Show()
        Me.Hide()

    End Sub
End Class
