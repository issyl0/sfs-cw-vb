Public Class frmReportMenu

    Private Sub btnQuit_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnQuit.Click
        'On close, close this window and show the main menu.
        Me.Close()
        frmMainMenu.Show()
    End Sub

    Private Sub btnQuotes_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnQuotes.Click
        'Show the quote form to the user, then hide this menu window.
        frmReportQuote.Show()
        Me.Hide()
    End Sub

    Private Sub btnLogs_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnLogs.Click
        'Show the log form to the user, then hide this menu window.
        frmReportLog.Show()
        Me.Hide()
    End Sub

    Private Sub btnInvoices_Click(ByVal sender As System.Object, _
                                  ByVal e As System.EventArgs) Handles btnInvoices.Click
        'Show the invoice form to the user, then hide this menu window.
        frmReportInvoice.Show()
        Me.Hide()
    End Sub

    Private Sub btnSurvey_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnSurvey.Click
        'Show the survey form to the user, then hide this menu window.
        frmReportSurvey.Show()
        Me.Hide()
    End Sub

End Class