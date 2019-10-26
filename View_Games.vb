' Name: Visual Basic Final Project 2019
' Purpose: To provide vital information about a collection of tabletop and video games
' Programmer: Derian Fritz (Started on: 4/28/2019)

Public Class frmViewGames

    Private Sub frmViewGames_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' This sub fills the Data Grid view with data from the Games table
        Me.GamesTableAdapter.FillByAll(Me.VB_Final_2019DataSet.Games)

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        ' When this button is clicked, the form is closed
        Me.Close()

    End Sub
End Class
