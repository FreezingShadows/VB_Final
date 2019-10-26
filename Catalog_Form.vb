' Name: Visual Basic Final Project 2019
' Purpose: To provide vital information about a collection of tabletop and video games
' Programmer: Derian Fritz (Started on: 4/28/2019)

Public Class frmCatalog

    ' This variable creates a new instance of the Log_Class
    Private logActivity As New Log_Class

    ' This variable establishes a new streamwriter for the log
    Private strLog As IO.StreamWriter

    Private Sub ReturnToMain(sender As Object, e As EventArgs) Handles ReturnToolStripMenuItem.Click, btnReturn.Click

        ' This sub handles both the return button, and the return tool strip item

        ' This variable creates a new instance of the main form
        Dim frmOne As New frmMain

        ' When thee items are pressed, the main form is shown and the current form is closed
        frmOne.Show()

        Me.Close()

    End Sub

    Private Sub ProgramExit(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click

        ' This sub handles the exit tool strip item

        ' When the exit item is pressed, the program is closed and a message is written to the log

        frmViewGames.Close()

        logActivity.EndOfSession(strLog, "")

        Me.Close()

    End Sub

    Private Sub frmCatalog_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' Fills the hidden data grid view with items from the Game_Id column
        Me.GamesTableAdapter1.Fill(Me.VB_Final_2019DataSetGames.Games)

        ' This Try...Catch block attempts to fill the data grid, and sends a message box if it fails
        Try
            Me.GamesTableAdapter.FillByGameId(Me.VB_Final_2019DataSet.Games)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub CreateCatalog(sender As Object, e As EventArgs) Handles CreateCatalogToolStripMenuItem.Click, btnCatalog.Click

        ' Declaring variables

        ' This variable counts the items selected from the list box
        Dim intCount As Integer

        ' This variable stores the number of selected values
        Dim intSelected As Integer

        ' This variable checks to see if the printer is created once
        Dim intHeaderCount As Integer = 0

        ' This variable stores the value chosen fron the list box as an integer
        Dim intVal As Integer

        ' This 2-dimensional array stores the header
        Dim strHeader(2, 1) As String

        ' This 2-dimensional array stores a row of values that will be output to the catalog
        Dim strDoc(4, 1) As String

        ' This variable holds text for the log
        Dim strText As String

        ' This variable establishes a streamwriter for the catalog
        Dim strNewCatalog As IO.StreamWriter

        ' Creates the catalog file
        strNewCatalog = IO.File.CreateText(txtFile.Text)

        ' These lines of code before the loop create the header values
        strHeader(0, 0) = "Name: "
        strHeader(1, 0) = "Date: "
        strHeader(2, 0) = "Number of Items: "

        strHeader(0, 1) = txtName.Text.Trim
        strHeader(1, 1) = DateTime.Now.ToString("MM/dd/yyyy")

        ' For the counter = 0 to the number of selected items minus 1, do these actions
        For intCount = 0 To lstColumns.SelectedItems.Count - 1
            If lstColumns.SelectedItems.Count() > 0 Then

                ' Grabs number of selected items
                intSelected = lstColumns.SelectedItems.Count()

                ' This while loop creates the catalog header
                Do While intHeaderCount = 0
                    strHeader(2, 1) = intSelected.ToString

                    strNewCatalog.Write(strHeader(0, 0))
                    strNewCatalog.WriteLine(strHeader(0, 1))
                    strNewCatalog.Write(strHeader(1, 0))
                    strNewCatalog.WriteLine(strHeader(1, 1))
                    strNewCatalog.Write(strHeader(2, 0))
                    strNewCatalog.WriteLine(strHeader(2, 1))
                    strNewCatalog.WriteLine("")

                    intHeaderCount = 1

                Loop

                ' This Try...Catch Block Attempts to store a selected value into a hidden label, outputs a message box if it fails

                Try
                    lblHidden.Text = lstColumns.SelectedItems(intCount)(lstColumns.ValueMember)
                Catch ex As Exception
                    MessageBox.Show("Error: Failure to store value in label..." & ex.Message, "Insert", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try

                ' This Try...Catch block attempts to grab the information of the chosen value in the database, and write it to a text file. A message box is output if it fails
                Try
                    Integer.TryParse(lblHidden.Text, intVal)
                    GamesTableAdapter1.FillById(VB_Final_2019DataSetGames.Games, intVal)

                    ' Static array values
                    strDoc(0, 0) = "Name: "
                    strDoc(1, 0) = "Developer: "
                    strDoc(2, 0) = "Number in stock: "
                    strDoc(3, 0) = "Price: $"
                    strDoc(4, 0) = "Description: "

                    ' Values from database
                    strDoc(0, 1) = GamesDataGridView.Rows(0).Cells(2).Value.ToString
                    strDoc(1, 1) = GamesDataGridView.Rows(0).Cells(3).Value.ToString
                    strDoc(2, 1) = GamesDataGridView.Rows(0).Cells(4).Value.ToString
                    strDoc(3, 1) = GamesDataGridView.Rows(0).Cells(6).Value.ToString
                    strDoc(4, 1) = GamesDataGridView.Rows(0).Cells(7).Value.ToString

                    ' Catalog being written
                    strNewCatalog.Write(strDoc(0, 0))
                    strNewCatalog.WriteLine(strDoc(0, 1))

                    strNewCatalog.Write(strDoc(1, 0))
                    strNewCatalog.WriteLine(strDoc(1, 1))

                    strNewCatalog.Write(strDoc(2, 0))
                    strNewCatalog.WriteLine(strDoc(2, 1))

                    strNewCatalog.Write(strDoc(3, 0))
                    strNewCatalog.WriteLine(strDoc(3, 1))

                    strNewCatalog.Write(strDoc(4, 0))
                    strNewCatalog.WriteLine(strDoc(4, 1))

                    strNewCatalog.WriteLine("")

                Catch ex As Exception
                    MessageBox.Show("Error: SQL Insert Query Failure..." & ex.Message, "Insert", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try

            End If
        Next

        ' Log text
        strText = "A Catalog titled " & txtFile.Text & " was created on: "

        ' Writes text and date to log
        logActivity.WriteToLog(strLog, strText)

        ' Closes the catalog writer
        strNewCatalog.Close()

        ' Outputs if catalog creation is successful
        MessageBox.Show("Catalog Successfully Created!", "Catalog", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub ViewGames(sender As Object, e As EventArgs) Handles ViewGamesToolStripMenuItem.Click

        ' This variable creates a new instance of the View Games form
        Dim frmView As New frmViewGames

        ' When this tool strip item is clicked, the View Games form is shown
        frmView.Show()

    End Sub

End Class
