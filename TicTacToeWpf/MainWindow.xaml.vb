Class MainWindow 
    Private strTurn As String = "X"
    Private Sub ButtonClicked(sender As Button, e As RoutedEventArgs)
        sender.Content = strTurn
        strTurn = If(strTurn = "X", "Y", "X")

        checkForEndOfGame()

    End Sub
    Dim btnGrid(,) As Button
    Private Sub checkForEndOfGame()
        btnGrid = {{btnX0, btnX1, btnX2}, {btnY0, btnY1, btnY2}, {btnZ0, btnZ1, btnZ2}}
        'check rows
        For intRow As Integer = 0 To 2
            Dim strFirstChar As String = btnGrid(intRow, 0).Content
            If strFirstChar.Length = 0 Then
                GoTo endRow
            End If
            For intCol As Integer = 1 To 2
                If strFirstChar = btnGrid(intRow, intCol).Content Then
                Else
                    GoTo endRow
                End If
            Next
            endGame(strFirstChar)
endRow:
        Next
        'check columns
        For intCol As Integer = 0 To 2
            Dim strFirstChar As String = btnGrid(0, intCol).Content
            If strFirstChar.Length = 0 Then
                GoTo endCol
            End If
            For intRow As Integer = 1 To 2
                If strFirstChar = btnGrid(intRow, intCol).Content Then
                Else
                    GoTo endCol
                End If
            Next
            endGame(strFirstChar)
endCol:
        Next
        'check diagonals
        If btnGrid(0, 0).Content.Length > 0 And btnGrid(0, 0).Content = btnGrid(1, 1).Content And btnGrid(1, 1).Content = btnGrid(2, 2).Content Then
            endGame(btnGrid(0, 0).Content)
        End If
        If btnGrid(0, 2).Content.Length > 0 And btnGrid(0, 2).Content = btnGrid(1, 1).Content And btnGrid(1, 1).Content = btnGrid(2, 0).Content Then
            endGame(btnGrid(0, 2).Content)
        End If

    End Sub
    Private Sub endGame(ByVal strWinner As String)
        MessageBox.Show(strWinner & " wins")
        For intRow As Integer = 0 To 2
            For intCol As Integer = 0 To 2
                btnGrid(intRow, intCol).Content = ""
            Next
        Next
    End Sub
End Class
