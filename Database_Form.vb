' Name: Visual Basic Final Project 2019
' Purpose: To provide vital information about a collection of tabletop and video games
' Programmer: Derian Fritz (Started on: 4/28/2019)

Public Class frmData

    ' Declaring Private Variables

    ' This variable is a new instance of the Text_Handling_Class
    Private textFunctions As New Text_Handling_Class

    ' This variable is a new instance of the Log_Class
    Private logActivity As New Log_Class

    ' This variable establishes a new StreamWriter for the log
    Private strLog As IO.StreamWriter

    ' This variable is an integer that checks if the button has been clicked, in order to run the select case statement
    Private intClickCheck As Integer = 0

    Private Sub GamesBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs)

        ' This Try...Catch block attempts to save data to the database, and displays a message box if it fails
        Try
            Me.Validate()
            Me.GamesBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(Me.VB_Final_2019DataSet)
        Catch ex As Exception
            MessageBox.Show("Error: Save Failed...", "Games", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

        ' This variable holds text for the log
        Dim strText1 As String

        ' Log text
        strText1 = "A change in the 'Games' table of the database was saved at: "

        ' Writes entry to the log
        logActivity.WriteToLog(strLog, strText1)

    End Sub

    Private Sub frmData_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' When the form loads, theis sub fills the Data Grid View and hides the main form
        Me.GamesTableAdapter.FillByAll(Me.VB_Final_2019DataSet.Games)

        frmMain.Hide()

    End Sub
    Private Sub btnCat_Click(sender As Object, e As EventArgs) Handles btnCat.Click

        ' When this button is clicked, the Category form is shown
        Dim frmCat As New frmCat

        frmCat.Show()

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        ' Declaring Variables

        ' This variable holds the category number
        Dim intCategory As Integer

        ' This variable holds the name of the game
        Dim strName As String

        ' This variable holds the company name
        Dim strCompany As String

        ' This variable holds the quantity of the item
        Dim intQuantity As Integer

        ' This variable holds the date
        Dim strDate As String

        ' This variable holds the price
        Dim dblPrice As Double

        ' Sets the click check calue to true
        intClickCheck = 1

        ' If the click check value is met, it runs this case statement
        ' If one of the radio buttons on the form are pressed, it searches the database for the item requested, and outputs a message box if nothing was entered
        Select Case intClickCheck = 1
            Case radAll.Checked
                GamesTableAdapter.FillByAll(VB_Final_2019DataSet.Games)
            Case radCategory.Checked
                If txtCategory.Text.Trim = String.Empty Then
                    MessageBox.Show("Please Enter a Category to search by!", "Games", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else

                    Integer.TryParse(txtCategory.Text.Trim, intCategory)
                    GamesTableAdapter.FillByCat(VB_Final_2019DataSet.Games, intCategory)
                End If
            Case radName.Checked
                If txtName.Text.Trim = String.Empty Then
                    MessageBox.Show("Please Enter a Name to search by!", "Games", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    strName = txtName.Text
                    GamesTableAdapter.FillBySpecificName(VB_Final_2019DataSet.Games, strName)
                End If
            Case radCompany.Checked
                If txtCompany.Text.Trim = String.Empty Then
                    MessageBox.Show("Please Enter a Company to search by!", "Games", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    strCompany = txtCompany.Text
                    GamesTableAdapter.FillByCompany(VB_Final_2019DataSet.Games, strCompany)
                End If
            Case radQuantity.Checked
                If txtQuantity.Text.Trim = String.Empty Then
                    MessageBox.Show("Please Enter a Quantity to search by!", "Games", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Integer.TryParse(txtQuantity.Text.Trim, intQuantity)
                    GamesTableAdapter.FillByAmount(VB_Final_2019DataSet.Games, intQuantity)
                End If
            Case radDate.Checked
                If txtDate.Text.Trim = String.Empty Then
                    MessageBox.Show("Please Enter a Company to search by!", "Games", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    strDate = txtDate.Text
                    GamesTableAdapter.FillByDate(VB_Final_2019DataSet.Games, strDate)
                End If
            Case radPrice.Checked
                If txtPrice.Text.Trim = String.Empty Then
                    MessageBox.Show("Please Enter a Quantity to search by!", "Games", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Double.TryParse(txtPrice.Text.Trim, dblPrice)
                    GamesTableAdapter.FillByPrice(VB_Final_2019DataSet.Games, dblPrice)
                End If
        End Select

        ' This variable holds text for the log
        Dim strText2 As String

        ' Log text
        strText2 = "A 'Games' table search was conducted at: "

        ' Writes text and date to log
        logActivity.WriteToLog(strLog, strText2)

        ' Sets the click check value back to zero
        intClickCheck = 0

    End Sub

    Private Sub btnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click

        ' This variable creates a new instance of the main form
        Dim frmOne As New frmMain

        ' When this button is pressed, the main form is shown and the current form is closed
        frmOne.Show()
        Me.Close()

    End Sub

    Private Sub txtCategory_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCategory.KeyPress

        ' This makes sure the values input in this text box are numbers only
        textFunctions.NumOnly(sender, e)

    End Sub

    Private Sub txtCompany_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCompany.KeyPress

        ' This makes sure the values input in this text box are letters only
        textFunctions.NumAndLettersOnly(sender, e)

    End Sub

    Private Sub txtName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtName.KeyPress

        ' This makes sure the values input in this text box are letters only
        textFunctions.NumAndLettersOnly(sender, e)

    End Sub

    Private Sub txtPrice_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrice.KeyPress

        ' This makes sure the values input in this text box are numbers and symbols
        textFunctions.NumAndSymbols(sender, e)

    End Sub

    Private Sub txtDate_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDate.KeyPress

        ' This makes sure the values input in this text box are numbers and symbols
        textFunctions.NumAndSymbols(sender, e)

    End Sub

    Private Sub txtQuantity_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuantity.KeyPress

        ' This makes sure the values input in this text box are numbers only
        textFunctions.NumOnly(sender, e)

    End Sub

    Private Sub BindingNavigatorDeleteItem_Click(sender As Object, e As EventArgs) Handles BindingNavigatorDeleteItem.Click

        ' This sub writes to the log when the delete item is clicked

        ' This variable stores the text for the log
        Dim strText3 As String

        ' Log text
        strText3 = "An entry from the 'Games' table was deleted at: "

        ' Writes text and date to log
        logActivity.WriteToLog(strLog, strText3)

    End Sub
End Class
