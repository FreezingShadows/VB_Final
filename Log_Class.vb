' Name: Visual Basic Final Project 2019
' Purpose: To provide vital information about a collection of tabletop and video games
' Programmer: Derian Fritz (Started on: 4/28/2019)

Public Class Log_Class

    ' Declaring private properties

    ' This property holds the current date and time
    Private Property currDate As String

    ' This property holds the document name
    Private Property docName As String

    Public Sub New()

        ' This sub is a constructor that sets values to the private properties

        currDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")

        docName = "activity.txt"

    End Sub

    Public Function WriteToLog(ByVal writer As IO.StreamWriter, ByVal text As String)

        ' If the log exists, the text being written is appended

        ' If not, a file is created, a message is written notifying this creation, then text is appended

        If IO.File.Exists(docName) Then

            writer = IO.File.AppendText(docName)

            writer.Write(text)
            writer.WriteLine(currDate)

            writer.Close()

        Else

            writer = IO.File.CreateText(docName)

            writer.Write("Created a new text file at: ")
            writer.WriteLine(currDate)
            writer.WriteLine("")

            writer.Write(text)
            writer.WriteLine(currDate)

            writer.Close()

        End If

    End Function

    Public Function EndOfSession(ByVal writer As IO.StreamWriter, ByVal break As String)

        ' A break is written to the file, signifying the end of a session

        writer = IO.File.AppendText(docName)

        writer.WriteLine(break)

        writer.Close()

    End Function

End Class
