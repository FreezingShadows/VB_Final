' Name: Visual Basic Final Project 2019
' Purpose: To provide vital information about a collection of tabletop and video games
' Programmer: Derian Fritz (Started on: 4/28/2019)


Public Class frmAdd

    ' Declaring Private Variables

    ' This variable declares a new instance of the Text_Handling_Class
    Private textFunctions As New Text_Handling_Class

    ' This variable declares a new instance of the Log_Class
    Private logActivity As New Log_Class

    ' This variable declares a new instance of StreamWriter for the log
    Private strLog As IO.StreamWriter

    Private Sub txtCatName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCatName.KeyPress

        ' This sub allows only letters to be entered into the text box

        textFunctions.LettersOnly(sender, e)

    End Sub

    Private Sub SelectForm(sender As Object, e As EventArgs) Handles radAddItem.Click, radImportItem.Click, radAddCat.Click

        ' This sub determines which group box becomes active depending on which radio button is selected

        If radAddItem.Checked Then
            grpCat.Enabled = False
            grpImport.Enabled = False
            grpAddItem.Enabled = True

        End If

        If radImportItem.Checked Then
            grpAddItem.Enabled = False
            grpCat.Enabled = False
            grpImport.Enabled = True
        End If

        If radAddCat.Checked Then
            grpAddItem.Enabled = False
            grpImport.Enabled = False
            grpCat.Enabled = True

        End If
    End Sub

    Private Sub frmAdd_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' When the form loads, the main form is hidden and the category and add item group boxes are disabled

        frmMain.Hide()

        grpCat.Enabled = False
        grpAddItem.Enabled = False

    End Sub

    Private Sub btnSubmitCat_Click(sender As Object, e As EventArgs) Handles btnSubmitCat.Click

        ' Declaring Variables

        ' This variable holds the category name that the user is entering
        Dim strCatName As String

        ' This variable holds the category type that the user selected from the combo box
        Dim strCatType As String

        ' Trimming inputs
        strCatName = txtCatName.Text.Trim
        strCatType = cmbCatType.Text.Trim

        ' This Try...Catch block attempts to insert the information the user has entered into the database, and throws back a message box if it fails
        Try
            CategoriesTableAdapter.InsertCat(strCatName, strCatType)
        Catch ex As Exception
            MessageBox.Show("Error: SQL Insert Query Failure..." & ex.Message, "Insert", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

        ' Saves changes
        Me.Validate()
        Me.CategoriesBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.VB_Final_2019DataSetCat)

        ' This variable holds the string that will be entered into the log
        Dim strText1 As String

        ' String for log
        strText1 = "Data was inserted into the 'Categories' table at: "

        ' Calls to class that writes message and date to log
        logActivity.WriteToLog(strLog, strText1)

        ' Message box to confirm category was added successfully
        MessageBox.Show("Saved Changes", "Insert Category", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub btnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click

        ' This variable represents a new instance of form main
        Dim frmOne As New frmMain

        ' When the return button is pressed, the main form appears and the current form closes
        frmOne.Show()
        Me.Close()

    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click

        ' Declaring Variables

        ' This variable is a new StreamReader variable for importing files
        Dim newFile As IO.StreamReader

        ' This variable will be used to hold the file name being imported
        Dim strFileName As String

        ' This array will hold each column value for the row being imported
        Dim straFileRead(7) As String

        ' Sets the file name to whatever is entered in the textbox
        strFileName = txtImport.Text

        ' This if statement checks if the file exists. If it does not, it throws an error in a message box
        If IO.File.Exists(strFileName) Then
            ' Opens the file with the name stored in strFileName
            newFile = IO.File.OpenText(strFileName)

            ' Loop will run until the end of the file is reached
            Do Until newFile.Peek = -1

                ' The following seven statements stores each column value in the respective array values
                straFileRead(0) = newFile.ReadLine
                straFileRead(1) = newFile.ReadLine
                straFileRead(2) = newFile.ReadLine
                straFileRead(3) = newFile.ReadLine
                straFileRead(4) = newFile.ReadLine
                straFileRead(5) = newFile.ReadLine
                straFileRead(6) = newFile.ReadLine

                ' This Try...Catch statement attempts to insert these values into the games table, and sends a message box if it fails
                Try
                    GamesTableAdapter.InsertArray(straFileRead(0), straFileRead(1), straFileRead(2), straFileRead(3), straFileRead(4), straFileRead(5), straFileRead(6))
                Catch ex As Exception
                    MessageBox.Show("Error: SQL Insert Query Failure..." & ex.Message, "Insert", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try

            Loop
        Else

            MessageBox.Show("Error: File cannot be found!", "Import File", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

        ' This variable holds the string that will be entered into the log
        Dim strText2 As String

        ' Log message
        strText2 = "Data was inserted into the 'Games' table via import at: "

        ' This writes the message and the date to the log
        logActivity.WriteToLog(strLog, strText2)

        ' This message box confirms the changes to the database have gone through
        MessageBox.Show("Saved Changes", "Insert Category", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub btnList_Click(sender As Object, e As EventArgs) Handles btnList.Click
        ' This variable is a new instance of the category form
        Dim frmCat As New frmCat

        ' When the list button is pressed, the category form appears
        frmCat.Show()
    End Sub

    Private Sub btnSubmitItem_Click(sender As Object, e As EventArgs) Handles btnSubmitItem.Click

        ' Declaring Variables

        ' This variable holds the category entered in the text box
        Dim strCategory As String

        ' This variable holds the game name entered in the text box
        Dim strName As String

        ' This variable holds the company name entered in the text box
        Dim strCompany As String

        ' This variable holds the date entered in the text box
        Dim strDate As String

        ' This variable holds the price entered in the text box
        Dim dblPrice As Double

        'This variable holds the quantity entered in the text box
        Dim intAmount As Integer

        ' This variable holds the new game's description
        Dim strDescription As String

        ' Converting input to the appropriate values
        strCategory = txtCategory.Text.Trim
        strName = txtGameName.Text.Trim
        strCompany = txtCompany.Text.Trim
        strDate = DateTime.Now.ToString("MM/dd/yyyy")
        strDescription = txtDescription.Text
        Double.TryParse(txtPrice.Text.Trim, dblPrice)
        Integer.TryParse(txtAmount.Text.Trim, intAmount)

        ' This Try...Catch block attempts to insert the input into the Games table, and displays a message box if it fails
        Try
            GamesTableAdapter.InsertGames(strCategory, strName, strCompany, intAmount, strDate, dblPrice, strDescription)
        Catch ex As Exception
            MessageBox.Show("Error: SQL Insert Query Failure..." & ex.Message, "Insert", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

        ' Saves changes
        Me.Validate()
        Me.GamesBindingSource.EndEdit()
        Me.TableAdapterManager1.UpdateAll(Me.VB_Final_2019DataSet)

        ' This variable holds the text for the log
        Dim strText3 As String

        ' Log text
        strText3 = "Data was inserted into the 'Games' table manually at: "

        ' Writes text and date to log
        logActivity.WriteToLog(strLog, strText3)

        ' Message box confirming changes to the database
        MessageBox.Show("Saved Changes", "Insert Item", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub txtCategory_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCategory.KeyPress

        ' This sub allows only numbers to be entered into the text box

        textFunctions.NumOnly(sender, e)

    End Sub

    Private Sub txtGameName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtGameName.KeyPress

        ' This sub allows only letters and Numbers to be entered into the text box

        textFunctions.NumAndLettersOnly(sender, e)

    End Sub

    Private Sub txtAmount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAmount.KeyPress

        ' This sub allows only numbers to be entered into the text box

        textFunctions.NumOnly(sender, e)

    End Sub

    Private Sub txtCompany_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCompany.KeyPress

        ' This sub allows only letters and Numbers to be entered into the text box

        textFunctions.NumAndLettersOnly(sender, e)

    End Sub

    Private Sub txtPrice_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrice.KeyPress

        ' This sub allows only numbers and symbols to be entered into the text box

        textFunctions.NumAndSymbols(sender, e)

    End Sub

End Class
