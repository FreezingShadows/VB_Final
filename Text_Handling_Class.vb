' Name: Visual Basic Final Project 2019
' Purpose: To provide vital information about a collection of tabletop and video games
' Programmer: Derian Fritz (Started on: 4/28/2019)

Public Class Text_Handling_Class
    Public Function NumOnly(sender As Object, e As KeyPressEventArgs)

        ' This function only allows numbers to be entered into a text box

        If Char.IsLetter(e.KeyChar) Then
            e.Handled = True

        End If

        If Char.IsPunctuation(e.KeyChar) Then
            e.Handled = True

        End If

        If Char.IsSeparator(e.KeyChar) Then
            e.Handled = True

        End If

        If Char.IsSymbol(e.KeyChar) Then
            e.Handled = True

        End If
    End Function
    Public Function NumAndSymbols(sender As Object, e As KeyPressEventArgs)

        ' This function allows only numbers and symbols like punctuation marks to be entered into a text box

        If Char.IsLetter(e.KeyChar) Then
            e.Handled = True

        End If

        If Char.IsSymbol(e.KeyChar) Then
            e.Handled = True

        End If

    End Function
    Public Function LettersAndSymbols(sender As Object, e As KeyPressEventArgs)

        ' This function allows only numbers and symbols like punctuation marks to be entered into a text box

        If Char.IsNumber(e.KeyChar) Then
            e.Handled = True

        End If

        If Char.IsSymbol(e.KeyChar) Then
            e.Handled = True

        End If

    End Function
    Public Function LettersOnly(sender As Object, e As KeyPressEventArgs)

        ' This function only allows numbers to be entered into a text box

        If Char.IsNumber(e.KeyChar) Then
            e.Handled = True

        End If

        If Char.IsPunctuation(e.KeyChar) Then
            e.Handled = True

        End If

        If Char.IsSymbol(e.KeyChar) Then
            e.Handled = True

        End If
    End Function
    Public Function NumAndLettersOnly(sender As Object, e As KeyPressEventArgs)

        ' This function only allows numbers and letters to be entered into a text box

        If Char.IsPunctuation(e.KeyChar) Then
            e.Handled = True

        End If

        If Char.IsSymbol(e.KeyChar) Then
            e.Handled = True

        End If

    End Function
End Class
