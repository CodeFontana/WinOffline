Imports System.Drawing
Imports System.Windows.Forms

Public Class AlertBox

    ' Class variables
    Private FormTitle As String
    Private AutoCloseCounter As Integer = 0
    Private AutoCloseTimeout As Integer = 0


    Public Shared Sub CreateUserAlert(ByVal Message As String,
                                      Optional ByVal Timeout As Integer = 0,
                                      Optional ByVal TextJustification As HorizontalAlignment = HorizontalAlignment.Center,
                                      Optional ByVal WordWrap As Boolean = True,
                                      Optional ByVal FontSize As Integer = 11,
                                      Optional ByVal AutoResizetoText As Boolean = False)

        ' Create a new instance
        Dim alert As New AlertBox()

        ' Set the titlebar
        alert.Text = Globals.ProcessFriendlyName + " " + Globals.AppVersion
        alert.FormTitle = alert.Text

        ' Set the message customizations
        alert.AlertText.Text = Message
        alert.AlertText.TextAlign = TextJustification
        alert.AlertText.WordWrap = WordWrap
        alert.AlertText.Font = New Font(alert.AlertText.Font.Name, FontSize, FontStyle.Regular)

        ' Check for auto resize to text option
        If AutoResizetoText Then

            ' Allow alert form to be auto sized
            alert.AutoSize = True
            alert.AutoSizeMode = AutoSizeMode.GrowAndShrink

            ' Measure rendered text size
            Dim Size As Size = TextRenderer.MeasureText(alert.AlertText.Text, alert.AlertText.Font)

            ' Adjust textbox dimensions
            alert.AlertText.Width = Size.Width

            ' Center button
            alert.btnOK.Location = New Point((alert.Width - alert.btnOK.Width) \ 2, alert.btnOK.Location.Y)

        End If

        ' Check for non-zero timeout
        If Timeout <> 0 Then

            ' Set the timeout
            alert.AutoCloseTimeout = Timeout

            ' Start the timer
            alert.Timer1.Start()

        End If

        ' Show modal dialog
        alert.ShowDialog()

    End Sub

    ' Timer tick
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        ' Check the interval
        If AutoCloseCounter >= AutoCloseTimeout Then

            ' Stop the timer
            Timer1.Stop()

            ' Close the form
            Me.Close()

        End If

        ' Increment the time
        AutoCloseCounter += 1

        ' Update the titlebar
        Text = FormTitle + " -- " + ((AutoCloseTimeout + 1) - AutoCloseCounter).ToString + " seconds."

    End Sub

    ' OK button pressed
    Private Sub OkayButton_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click

        ' Close the form
        Close()

    End Sub

End Class