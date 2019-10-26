' Name: Visual Basic Final Project 2019
' Purpose: To provide vital information about a collection of tabletop and video games
' Programmer: Derian Fritz (Started on: 4/28/2019)

Public Class frmManageCat

    ' Declaring private variables

    ' This variable declares a new instance of the log class
    Private logActivity As New Log_Class

    ' This variable creates a new StreamWriter for the log
    Private strLog As IO.StreamWriter

    Private Sub frmManageCat_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' When the form is loaded, the data grid view is loaded and the main form is hidden
        Me.CategoriesTableAdapter.Fill(Me.VB_Final_2019DataSetCat.Categories)
        frmMain.Hide()

    End Sub

    Private Sub CategoriesBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs) Handles CategoriesBindingNavigatorSaveItem.Click

        ' This variable stores text for the log
        Dim strText As String

        ' Changes saved
        Me.Validate()
        Me.CategoriesBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.VB_Final_2019DataSetCat)

        ' Log text
        strText = "A change in the 'Categories' table of the database was saved at: "

        ' Writes text and date to log
        logActivity.WriteToLog(strLog, strText)


    End Sub

    Private Sub btnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click
        ' This variable creates a new instance of form main
        Dim frmOne As New frmMain

        ' The main form is shown, and the current form is closed
        frmOne.Show()
        Me.Close()

    End Sub

    Private Sub BindingNavigatorDeleteItem_Click(sender As Object, e As EventArgs) Handles BindingNavigatorDeleteItem.Click

        ' This variable stores text for the log
        Dim strText As String

        ' Log text
        strText = "An entry from the 'Categories' table was deleted at: "

        ' Writes text and date to log
        logActivity.WriteToLog(strLog, strText)

    End Sub
End Class
