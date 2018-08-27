Public Class HelpUI

    Private Sub btnReturn_Click(sender As Object, e As EventArgs) Handles btnReturn.Click
        Me.Close()
    End Sub

    Private Sub OptionsGui_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        ' Local variables
        Dim rtbText As New ArrayList

        ' Replace WinOffline references with process name
        rtbHelp.Text = WinOffline.Utility.ReplaceWord(rtbHelp.Text, "WinOffline", Globals.ProcessFriendlyName)

        ' Iterate text from command line switches text box
        For Each line As String In rtbHelp.Lines

            ' Add each line to an array for processing
            rtbText.Add(line)

        Next

        ' Clear the text box
        rtbHelp.Text = ""

        ' Iterate the array
        For Each line As String In rtbText

            ' Append a line of text, from the array, to the text box
            rtbHelp.AppendText(line + Environment.NewLine)

            ' Check for special formatting
            If line.Contains(Globals.ProcessFriendlyName + " Command Line Switches:") Or
                line.Contains(Globals.ProcessFriendlyName + " Software Delivery and Scripting Switches:") Or
                line.Contains(Globals.ProcessFriendlyName + " Database Maintenance Switches:") Then

                ' Select the text
                rtbHelp.SelectionStart = rtbHelp.Text.LastIndexOf(line)
                rtbHelp.SelectionLength = line.Length

                ' Update formatting
                rtbHelp.SelectionFont = New Drawing.Font("Calibri", 12, Drawing.FontStyle.Underline Or Drawing.FontStyle.Bold)

            End If

        Next

        ' Deselect text
        rtbHelp.SelectionStart = 0
        rtbHelp.SelectionLength = 0

    End Sub

End Class