Public NotInheritable Class AboutBox

    Private Sub AboutBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Set titlebar text
        Text = String.Format("About {0}", Globals.ProcessFriendlyName)

        ' Set product name text
        LabelProductName.Text = Globals.ProcessFriendlyName

        ' Set product version text
        LabelVersion.Text = String.Format("Version {0}", Globals.AppVersion)

        ' Set copyright text
        LabelCopyright.Text = "Copyright " + Globals.AppVersion.Substring(0, Globals.AppVersion.IndexOf("."))

        ' Set company name text
        LabelCompanyName.Text = "CA Technologies, Inc."

    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ' Close the form
        Close()

    End Sub

End Class