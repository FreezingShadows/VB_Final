' Name: Visual Basic Final Project 2019
' Purpose: To provide vital information about a collection of tabletop and video games
' Programmer: Derian Fritz (Started on: 4/28/2019)

Public Class frmCat

    Private Sub Category_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' This sub fills the table on this form with data from the categories table

        Me.CategoriesTableAdapter.Fill(Me.VB_Final_2019DataSetCat.Categories)

    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        ' When this button is clicked, the form is closed
        Me.Close()

    End Sub
End Class
