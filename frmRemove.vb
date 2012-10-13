Public Class frmRemove

    Private Sub frmRemove_Load(ByVal sender As System.Object, _
                               ByVal e As System.EventArgs) Handles MyBase.Load
        'Check if the database connection is breathing.
        'If it isn't, resuscitate it.  :-)
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        cbtxtRemoveOptions.Items.Add("Customer")
        cbtxtRemoveOptions.Items.Add("Supplier")
        cbtxtRemoveOptions.Items.Add("Component")
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnShow.Click

        'According to the option selected, display the next form and hide
        'this one.
        If cbtxtRemoveOptions.Text = "Customer" Then
            Me.Hide()
            frmRemoveCustomer.Show()
        ElseIf cbtxtRemoveOptions.Text = "Supplier" Then
            Me.Hide()
            frmRemoveSupplier.Show()
        ElseIf cbtxtRemoveOptions.Text = "Component" Then
            Me.Hide()
            frmRemoveComponent.Show()
        End If

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, _
                               ByVal e As System.EventArgs) Handles btnClose.Click
        'Close this window and show the main menu.
        Me.Close()
        frmMainMenu.Show()
    End Sub
End Class