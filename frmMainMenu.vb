Imports System.Data
Imports System.Data.OleDb

Public Class frmMainMenu
    Private Sub frmMainMenu_Load(ByVal sender As System.Object, _
                                 ByVal e As System.EventArgs) Handles MyBase.Load
        'Hide the login form from the user when this form is loaded so
        'that he/she is not having to close a load of windows when he/she
        'quits the program.
        frmLoginForm.Hide()
    End Sub

    'In each of these, according to the buttons clicked, show the correct
    'form and close this window.
    Private Sub btnReport_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnReport.Click
        Me.Hide()
        frmReportMenu.Show()
    End Sub

    Private Sub btnList_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnList.Click
        Me.Hide()
        frmList.Show()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, _
                             ByVal e As System.EventArgs) Handles btnAdd.Click
        Me.Hide()
        frmAdd.Show()
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnRemove.Click
        Me.Hide()
        frmRemove.Show()
    End Sub

    Private Sub btnQuit_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnQuit.Click
        Me.Close()
        'Now close the login form that is at present hidden so that the
        'program closes and doesn't stay running displaying nothing.  Also,
        'someone would have already logged in, so displaying the login
        'form again would intice people into logging in again, potentially
        'messing database stuff up like the closing of connections etc.
        frmLoginForm.Close()
    End Sub

End Class
