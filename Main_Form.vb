' Name: Visual Basic Final Project 2019
' Purpose: To provide vital information about a collection of tabletop and video games
' Programmer: Derian Fritz (Started on: 4/28/2019)


Public Class frmMain

    ' Declaring Private Variables

    ' This variable creates a new instance of the Log_Class
    Private logActivity As New Log_Class

    ' This variable establishes a new StreamWriter for the log
    Private strLog As IO.StreamWriter

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' When the form is loaded, all other forms are hidden
        frmData.Hide()
        frmAdd.Hide()
        frmCatalog.Hide()
        frmManageCat.Hide()

    End Sub
    Private Sub btnView_Click(sender As Object, e As EventArgs) Handles btnView.Click

        ' This variable creates a new instance of the database form
        Dim frmTwo As New frmData

        ' The database form is shown
        frmTwo.Show()

        ' This variable holds text for the log
        Dim strText1 As String

        ' Log text
        strText1 = "The 'Manage Items' form was accessed at: "

        ' The text and date are written to the log
        logActivity.WriteToLog(strLog, strText1)

        ' Main form is closed
        Me.Close()

    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click

        ' This variable creates a new instance of the addition form
        Dim frmThree As New frmAdd

        ' Addition form is shown
        frmThree.Show()

        ' This variable holds text for the log
        Dim strText2 As String

        ' Log text
        strText2 = "The 'New Entry' form was accessed at: "

        ' The text and date are written to the log
        logActivity.WriteToLog(strLog, strText2)

        ' Main form is closed
        Me.Close()

    End Sub
    Private Sub btnCatalog_Click(sender As Object, e As EventArgs) Handles btnCatalog.Click

        ' This variable creates a new instance of the catalog form
        Dim frmFour As New frmCatalog

        ' Catalog form is shown
        frmFour.Show()

        ' This variable holds text for the log
        Dim strText3 As String

        ' Log text
        strText3 = "The 'Print Catalog' form was accessed at: "

        ' The text and date are written to the log
        logActivity.WriteToLog(strLog, strText3)

        ' Main form is closed
        Me.Close()

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        ' This variable holds the text for the log
        Dim strText4 As String

        ' Log text
        strText4 = "The program was closed at: "

        ' The text and date are written to the log
        logActivity.WriteToLog(strLog, strText4)

        ' A break is written to the log
        logActivity.EndOfSession(strLog, "")

        ' Main form is closed
        Me.Close()

    End Sub

    Private Sub btnCat_Click(sender As Object, e As EventArgs) Handles btnCat.Click

        ' This variable creates a new instance of the manage categories form
        Dim frmFive As New frmManageCat

        ' Manage categories form is shown
        frmFive.Show()

        ' This variable holds text for the log
        Dim strText5 As String

        ' Log text
        strText5 = "The 'Manage Catagories' form was accessed at: "

        ' The text and date are written to the log
        logActivity.WriteToLog(strLog, strText5)

        ' Main form is closed
        Me.Close()

    End Sub

    Private Sub btnView_MouseEnter(sender As Object, e As EventArgs) Handles btnView.MouseEnter

        ' Image is shown in picture box when this button has the focus of the mouse
        picLogo.Image = Image.FromFile("search icon.png")

    End Sub

    Private Sub btnCat_MouseEnter(sender As Object, e As EventArgs) Handles btnCat.MouseEnter

        ' Image is shown in picture box when this button has the focus of the mouse
        picLogo.Image = Image.FromFile("folder icon.png")

    End Sub

    Private Sub btnNew_MouseEnter(sender As Object, e As EventArgs) Handles btnNew.MouseEnter

        ' Image is shown in picture box when this button has the focus of the mouse
        picLogo.Image = Image.FromFile("box icon.png")

    End Sub

    Private Sub btnCatalog_MouseEnter(sender As Object, e As EventArgs) Handles btnCatalog.MouseEnter

        ' Image is shown in picture box when this button has the focus of the mouse
        picLogo.Image = Image.FromFile("print icon.png")

    End Sub

    Private Sub btnExit_MouseEnter(sender As Object, e As EventArgs) Handles btnExit.MouseEnter

        ' Image is shown in picture box when this button has the focus of the mouse
        picLogo.Image = Image.FromFile("exit icon.png")

    End Sub

    Private Sub ReturnImage(sender As Object, e As EventArgs) Handles btnView.MouseLeave, btnCat.MouseLeave, btnNew.MouseLeave, btnCatalog.MouseLeave, btnExit.MouseLeave

        ' Image is shown in picture box when no button has the focus of the mouse
        picLogo.Image = Image.FromFile("safe icon.png")

    End Sub
End Class
