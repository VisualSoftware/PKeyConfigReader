Public Class VSSharedSource
    Dim donationpage As String = "https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=XTCNPRUDFH4UA&lc=ES&item_name=PKEYCONFIG&currency_code=EUR&bn=PP%2dDonationsBF%3abtn_donate_LG%2egif%3aNonHosted"
    Dim appname As String = "PKEYCONFIG"

    Function OpenDonationPage()
        Try
            Process.Start(donationpage)
        Catch ex As Exception
            MessageBox.Show("Error: The Donation page couldn't be opened. Open your favorite browser and go to: " & donationpage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Function

    Function GetCopyrightDate() As String
        Dim cpdate As String
        If Now.Year > 2013 Then
            cpdate = " © 2013-" & Now.Year
        Else
            cpdate = " © 2013"
        End If
        Return cpdate
    End Function

End Class
