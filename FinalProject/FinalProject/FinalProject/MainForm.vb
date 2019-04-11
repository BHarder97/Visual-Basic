Public Class MainForm
    'Brandon Harder
    'Computer Science
    'David DeVary
    'Final Project
    'If you click too fast you will break the game.

    Dim stopwatch As New Stopwatch
    Dim Var1 As New PictureBox
    Dim Var2 As New PictureBox

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Calls the function to assign each picturebox an image
        AssigningPictures()

        'Calls the function to shuffle the cards around
        Shuffle()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Closes the form
        Me.Close()
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        'Resets the form.  Does not reset the time.  If you want the timer to reset you have to quit then start again.
        HideCards()
        Timer1.Enabled = False
        Timer1.Stop()
        Timer1.Enabled = True
        Timer1.Enabled = False
        Me.stopwatch.Reset()
        Label2.Text = "00:00:00:000"
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click, PictureBox2.Click, PictureBox3.Click, PictureBox4.Click, PictureBox5.Click, PictureBox6.Click, PictureBox7.Click, PictureBox8.Click, PictureBox9.Click, PictureBox10.Click, PictureBox11.Click, PictureBox12.Click, PictureBox13.Click, PictureBox14.Click, PictureBox15.Click, PictureBox16.Click, PictureBox17.Click, PictureBox18.Click, PictureBox19.Click, PictureBox20.Click, PictureBox21.Click, PictureBox22.Click, PictureBox23.Click, PictureBox24.Click, PictureBox25.Click, PictureBox26.Click, PictureBox27.Click, PictureBox28.Click, PictureBox29.Click, PictureBox30.Click, PictureBox31.Click, PictureBox32.Click, PictureBox33.Click, PictureBox34.Click, PictureBox35.Click, PictureBox36.Click, PictureBox37.Click, PictureBox38.Click, PictureBox39.Click, PictureBox40.Click

        Dim tempPbx As New PictureBox               ' this will be the current picture box that is clicked
        Static firstpb As New PictureBox            ' this will be first picture box that is clicked

        Static Counter As Integer = 0    'Coutner for matches

        Static clickCount As Integer = 1            ' static because, i dont' want this varible to deleted
        '                                             at the end of the click sub

        stsLabel.Text = ""

        tempPbx = CType(sender, PictureBox)         'convert the parameter into a picturebox

        If clickCount = 1 Then                      ' only executes if the first click. 
            firstpb = tempPbx                       ' saves which picture box was clicked first
        End If

        'Checks to see what the tag is and then assigns the picturebox the correct image.
        If tempPbx.Tag = "Grass" Then
            tempPbx.Image = Pictures(0)

        ElseIf tempPbx.Tag = "Fire" Then
            tempPbx.Image = Pictures(1)

        ElseIf tempPbx.Tag = "Water" Then
            tempPbx.Image = Pictures(2)

        ElseIf tempPbx.Tag = "Lightning" Then
            tempPbx.Image = Pictures(3)

        ElseIf tempPbx.Tag = "Lugia" Then
            tempPbx.Image = Pictures(4)

        ElseIf tempPbx.Tag = "Moltres" Then
            tempPbx.Image = Pictures(5)

        ElseIf tempPbx.Tag = "Zapdos" Then
            tempPbx.Image = Pictures(6)

        ElseIf tempPbx.Tag = "Articuno" Then
            tempPbx.Image = Pictures(7)

        ElseIf tempPbx.Tag = "Entei" Then
            tempPbx.Image = Pictures(8)

        ElseIf tempPbx.Tag = "Suicune" Then
            tempPbx.Image = Pictures(9)

        ElseIf tempPbx.Tag = "Raikou" Then
            tempPbx.Image = Pictures(10)

        ElseIf tempPbx.Tag = "Charmander" Then
            tempPbx.Image = Pictures(11)

        ElseIf tempPbx.Tag = "Squirtle" Then
            tempPbx.Image = Pictures(12)

        ElseIf tempPbx.Tag = "Bulbasaur" Then
            tempPbx.Image = Pictures(13)

        ElseIf tempPbx.Tag = "Groudon" Then
            tempPbx.Image = Pictures(14)

        ElseIf tempPbx.Tag = "Kyogre" Then
            tempPbx.Image = Pictures(15)

        ElseIf tempPbx.Tag = "Rayquaza" Then
            tempPbx.Image = Pictures(16)

        ElseIf tempPbx.Tag = "Regirock" Then
            tempPbx.Image = Pictures(17)

        ElseIf tempPbx.Tag = "Regice" Then
            tempPbx.Image = Pictures(18)

        ElseIf tempPbx.Tag = "Registeel" Then
            tempPbx.Image = Pictures(19)
        End If


        'If the tags match and its not the first click
        If clickCount = 2 Then
            If firstpb.Name = tempPbx.Name Then
                Exit Sub
            End If
            If firstpb.Tag = tempPbx.Tag Then
                stsLabel.Text = "Tags Matched!  You found a pair!!"
                tempPbx.Visible = False
                firstpb.Visible = False
                Counter += 1
            Else                                    ' tags don't match on 2nd click
                stsLabel.Text = "Tags did not match.  Sorry."
                Var1 = tempPbx
                Var2 = firstpb
                Timer2.Enabled = True
            End If
            clickCount = 1                          ' resets count back to one after 2nd click
        Else
            clickCount += 1                         ' if 1st click add one to count
        End If

        If Counter = 20 Then
            stsLabel.Text = "You Won!!"
            Timer1.Enabled = False
            Timer1.Stop()
            MessageBox.Show("CONGRATULATIONS YOU HAVE WON!")

        End If

    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        'Calling to show the cards and then calls to shuffle them
        HideCards()
        ShowCards()
        Shuffle()

        'Starts the timer
        Timer1.Enabled = True
        Timer1.Start()
        Me.stopwatch.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'The code for the timer
        Dim elapsed As TimeSpan = Me.stopwatch.Elapsed
        Label2.Text = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", _
        Math.Floor(elapsed.TotalHours), _
        elapsed.Minutes, elapsed.Seconds, _
        elapsed.Milliseconds)
    End Sub

    Sub HideCards()
        'Making all the pictureboxes have the pokemon card as the image
        For i As Integer = 0 To 39
            PicBoxes(i).Image = My.Resources.PokemonBack
        Next

        'Making all the cards not show
        For i As Integer = 0 To 39
            PicBoxes(i).Visible = False
        Next
    End Sub

    Sub AssigningPictures()
        'Assigns each picturebox to a picture
        For i = 0 To 39
            PicBoxes(i).Tag = picTags(i)
        Next
    End Sub

    Sub Shuffle()
        'Creates some variables for the shuffle loop
        Dim Holder As New PictureBox
        Dim Holder2 As Integer

        Dim rand As New Random

        'Shuffles the pictureboxes so that they are in different locations
        For i As Integer = 0 To 39

            Holder2 = rand.Next(40)
            Holder.Location = PicBoxes(i).Location
            PicBoxes(i).Location = PicBoxes(Holder2).Location
            PicBoxes(Holder2).Location = Holder.Location

        Next
    End Sub

    Sub ShowCards()
        'Makes all the cards visible
        For i = 0 To 39
            PicBoxes(i).Visible = True
        Next
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Var1.Image = My.Resources.PokemonBack
        Var2.Image = My.Resources.PokemonBack
        Timer2.Enabled = False
    End Sub
End Class
