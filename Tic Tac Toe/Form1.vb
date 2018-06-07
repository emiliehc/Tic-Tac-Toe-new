'Frank Chen
'mar 28
' tic tac toe game










Public Class Form1

    Dim xcase1, xcase2, xcase3, xcase4, xcase5, xcase6, xcase7, xcase8 As Boolean
    Dim ocase1, ocase2, ocase3, ocase4, ocase5, ocase6, ocase7, ocase8 As Boolean
    Dim tie As Boolean
    Dim gameStarted As Boolean = True
    Dim done, blocked As Boolean
    Dim number As Integer
    Dim gameBoard(0 To 8) As Button


    Sub readGameBoard()
        gameBoard(0) = btn0
        gameBoard(1) = btn1
        gameBoard(2) = btn2
        gameBoard(3) = btn3
        gameBoard(4) = btn4
        gameBoard(5) = btn5
        gameBoard(6) = btn6
        gameBoard(7) = btn7
        gameBoard(8) = btn8
    End Sub

    Sub applyGameBoard()
        btn0 = gameBoard(0)
        btn1 = gameBoard(1)
        btn2 = gameBoard(2)
        btn3 = gameBoard(3)
        btn4 = gameBoard(4)
        btn5 = gameBoard(5)
        btn6 = gameBoard(6)
        btn7 = gameBoard(7)
        btn8 = gameBoard(8)
    End Sub

    Sub generateNextStep()
        If gameStarted = True Then
            done = False ' Make sure that the computer generates an "O":
            blocked = False
            ' first find if it is possible to block the user aka finding
            ' 2 X's in a row and place an O on the third cell
            If blocked = False Then

                ' If possible, occupy the center
                If gameBoard(4).Enabled = True Then
                    With gameBoard(4)
                        .Text = "O"
                        .Enabled = False
                    End With
                    blocked = True
                End If


                ' If X occupies the center ...
                If gameBoard(4).Enabled = False And gameBoard(4).Text = "X" And gameBoard(2).Enabled = True And blocked = False Then
                    With gameBoard(2)
                        .Text = "O"
                        .Enabled = False
                    End With
                    blocked = True
                End If

                ' check horizontal
                For x = 0 To 6 Step 3
                    If gameBoard(x).Text = gameBoard(x + 1).Text And gameBoard(x).Text <> "" And gameBoard(x + 2).Enabled = True And blocked = False Then
                        With gameBoard(x + 2)
                            .Text = "O"
                            .Enabled = False
                        End With
                        blocked = True
                    ElseIf gameBoard(x + 1).Text = gameBoard(x + 2).Text And gameBoard(x + 1).Text <> "" And gameBoard(x).Enabled = True And blocked = False Then
                        With gameBoard(x)
                            .Text = "O"
                            .Enabled = False
                        End With
                        blocked = True
                    ElseIf gameBoard(x).Text = gameBoard(x + 2).Text And gameBoard(x).Text <> "" And gameBoard(x + 1).Enabled = True And blocked = False Then
                        With gameBoard(x + 1)
                            .Text = "O"
                            .Enabled = False
                        End With
                        blocked = True
                    End If
                Next

                'check vertical
                For x = 0 To 2
                    If gameBoard(x).Text = gameBoard(x + 3).Text And gameBoard(x).Text <> "" And gameBoard(x + 6).Enabled = True And blocked = False Then
                        With gameBoard(x + 6)
                            .Text = "O"
                            .Enabled = False
                        End With
                        blocked = True
                    ElseIf gameBoard(x + 3).Text = gameBoard(x + 6).Text And gameBoard(x + 3).Text <> "" And gameBoard(x).Enabled = True And blocked = False Then
                        With gameBoard(x)
                            .Text = "O"
                            .Enabled = False
                        End With
                        blocked = True
                    ElseIf gameBoard(x).Text = gameBoard(x + 6).Text And gameBoard(x).Text <> "" And gameBoard(x + 3).Enabled = True And blocked = False Then
                        With gameBoard(x + 3)
                            .Text = "O"
                            .Enabled = False
                        End With
                        blocked = True
                    End If
                Next


                ' check diagnal
                If gameBoard(0).Text = gameBoard(4).Text And gameBoard(0).Text <> "" And gameBoard(8).Enabled = True And blocked = False Then
                    With gameBoard(8)
                        .Text = "O"
                        .Enabled = False
                    End With
                    blocked = True
                ElseIf gameBoard(0).Text = gameBoard(8).Text And gameBoard(0).Text <> "" And gameBoard(4).Enabled = True And blocked = False Then
                    With gameBoard(4)
                        .Text = "O"
                        .Enabled = False
                    End With
                    blocked = True
                ElseIf gameBoard(4).Text = gameBoard(8).Text And gameBoard(4).Text <> "" And gameBoard(0).Enabled = True And blocked = False Then
                    With gameBoard(0)
                        .Text = "O"
                        .Enabled = False
                    End With
                    blocked = True
                End If

                If gameBoard(2).Text = gameBoard(4).Text And gameBoard(2).Text <> "" And gameBoard(6).Enabled = True And blocked = False Then
                    With gameBoard(6)
                        .Text = "O"
                        .Enabled = False
                    End With
                    blocked = True
                ElseIf gameBoard(2).Text = gameBoard(6).Text And gameBoard(2).Text <> "" And gameBoard(4).Enabled = True And blocked = False Then
                    With gameBoard(4)
                        .Text = "O"
                        .Enabled = False
                    End With
                    blocked = True
                ElseIf gameBoard(4).Text = gameBoard(6).Text And gameBoard(4).Text <> "" And gameBoard(2).Enabled = True And blocked = False Then
                    With gameBoard(2)
                        .Text = "O"
                        .Enabled = False
                    End With
                    blocked = True
                End If




                ' if the computer fails to block, generate a cell randomly
                If blocked = False Then
                    Do While done = False
                        Randomize()
                        ' generate a random number
                        number = Int(Rnd() * 9) ' choose one randomly 

                        If gameBoard(number).Enabled = True Then
                            With gameBoard(number)
                                .Enabled = False
                                .Text = "O"
                            End With
                            done = True
                        End If
                    Loop
                End If
            End If
        End If
    End Sub

    Sub checkWinner()
        applyGameBoard()

        ' Check for X

        xcase1 = gameBoard(0).Text = "X" And gameBoard(3).Text = "X" And gameBoard(6).Text = "X"
        xcase2 = gameBoard(1).Text = "X" And gameBoard(4).Text = "X" And gameBoard(7).Text = "X"
        xcase3 = gameBoard(2).Text = "X" And gameBoard(5).Text = "X" And gameBoard(8).Text = "X"
        xcase4 = gameBoard(0).Text = "X" And gameBoard(1).Text = "X" And gameBoard(2).Text = "X"
        xcase5 = gameBoard(3).Text = "X" And gameBoard(4).Text = "X" And gameBoard(5).Text = "X"
        xcase6 = gameBoard(6).Text = "X" And gameBoard(7).Text = "X" And gameBoard(8).Text = "X"
        xcase7 = gameBoard(0).Text = "X" And gameBoard(4).Text = "X" And gameBoard(8).Text = "X"
        xcase8 = gameBoard(6).Text = "X" And gameBoard(4).Text = "X" And gameBoard(2).Text = "X"

        ' Check for O
        ocase1 = gameBoard(0).Text = "O" And gameBoard(3).Text = "O" And gameBoard(6).Text = "O"
        ocase2 = gameBoard(1).Text = "O" And gameBoard(4).Text = "O" And gameBoard(7).Text = "O"
        ocase3 = gameBoard(2).Text = "O" And gameBoard(5).Text = "O" And gameBoard(8).Text = "O"
        ocase4 = gameBoard(0).Text = "O" And gameBoard(1).Text = "O" And gameBoard(2).Text = "O"
        ocase5 = gameBoard(3).Text = "O" And gameBoard(4).Text = "O" And gameBoard(5).Text = "O"
        ocase6 = gameBoard(6).Text = "O" And gameBoard(7).Text = "O" And gameBoard(8).Text = "O"
        ocase7 = gameBoard(0).Text = "O" And gameBoard(4).Text = "O" And gameBoard(8).Text = "O"
        ocase8 = gameBoard(6).Text = "O" And gameBoard(4).Text = "O" And gameBoard(2).Text = "O"

        'check tie
        tie = (gameBoard(0).Text <> Nothing) And (gameBoard(3).Text <> Nothing) And (gameBoard(6).Text <> Nothing) And (gameBoard(1).Text <> Nothing) And (gameBoard(4).Text <> Nothing) And (gameBoard(7).Text <> Nothing) And (gameBoard(2).Text <> Nothing) And (gameBoard(5).Text <> Nothing) And (gameBoard(8).Text <> Nothing)

        ' Check who wins
        If xcase1 Or xcase2 Or xcase3 Or xcase4 Or xcase5 Or xcase6 Or xcase7 Or xcase8 Then
            For x = 0 To 8
                gameBoard(x).Enabled = False
            Next
            lblShow.Text = "X wins"
            gameStarted = False
        ElseIf ocase1 Or ocase2 Or ocase3 Or ocase4 Or ocase5 Or ocase6 Or ocase7 Or ocase8 Then
            For x = 0 To 8
                gameBoard(x).Enabled = False
            Next
            lblShow.Text = "O wins"
            gameStarted = False
        ElseIf tie = True Then
            For x = 0 To 8
                gameBoard(x).Enabled = False
            Next
            lblShow.Text = "Draw"
            gameStarted = False
        End If
    End Sub


    Sub computerNextStep()
        If gameStarted = True Then
            done = False

        End If
    End Sub


    Private Sub btn0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn0.Click
        gameBoard(0).Text = "X"
        gameBoard(0).Enabled = False
        checkWinner()
        generateNextStep()
        applyGameBoard()
        checkWinner()
    End Sub

    Private Sub btn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn3.Click
        gameBoard(3).Text = "X"
        gameBoard(3).Enabled = False
        checkWinner()
        generateNextStep()
        applyGameBoard()
        checkWinner()
    End Sub

    Private Sub btn6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn6.Click
        gameBoard(6).Text = "X"
        gameBoard(6).Enabled = False
        checkWinner()
        generateNextStep()
        applyGameBoard()
        checkWinner()
    End Sub

    Private Sub btn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1.Click
        gameBoard(1).Text = "X"
        gameBoard(1).Enabled = False
        checkWinner()
        generateNextStep()
        applyGameBoard()
        checkWinner()
    End Sub

    Private Sub btn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn4.Click
        gameBoard(4).Text = "X"
        gameBoard(4).Enabled = False
        checkWinner()
        generateNextStep()
        applyGameBoard()
        checkWinner()
    End Sub

    Private Sub btn7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn7.Click
        gameBoard(7).Text = "X"
        gameBoard(7).Enabled = False
        checkWinner()
        generateNextStep()
        applyGameBoard()
        checkWinner()
    End Sub

    Private Sub btn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2.Click
        gameBoard(2).Text = "X"
        gameBoard(2).Enabled = False
        checkWinner()
        generateNextStep()
        applyGameBoard()
        checkWinner()
    End Sub

    Private Sub btn5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn5.Click
        gameBoard(5).Text = "X"
        gameBoard(5).Enabled = False
        checkWinner()
        generateNextStep()
        applyGameBoard()
        checkWinner()
    End Sub

    Private Sub btn8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn8.Click
        gameBoard(8).Text = "X"
        gameBoard(8).Enabled = False
        checkWinner()
        generateNextStep()
        applyGameBoard()
        checkWinner()
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        For x = 0 To 8
            With gameBoard(x)
                .Enabled = True
                .Text = ""
            End With
        Next
        gameStarted = True
        lblShow.Text = ""
        readGameBoard()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        readGameBoard()
    End Sub



End Class
