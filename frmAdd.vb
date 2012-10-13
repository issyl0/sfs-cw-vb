Public Class frmAdd

    Private Sub frmAdd_Load(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs) Handles MyBase.Load
        'Check if the database connection is breathing.
        'If it isn't, resuscitate it.  :)
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        'Populate drop down box.
        cbtxtAddOptions.Items.Add("Customer")
        cbtxtAddOptions.Items.Add("Supplier")
        cbtxtAddOptions.Items.Add("Component")
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnShow.Click

        'According to the option selected, display the next form and hide
        'this one.
        If cbtxtAddOptions.Text = "Customer" Then
            Me.Hide()
            frmAddCustomer.Show()
        ElseIf cbtxtAddOptions.Text = "Supplier" Then
            Me.Hide()
            frmAddSupplier.Show()
        ElseIf cbtxtAddOptions.Text = "Component" Then
            Me.Hide()
            frmAddComponent.Show()
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, _
                               ByVal e As System.EventArgs) Handles btnClose.Click
        'Close this window and show the main menu.
        Me.Close()
        frmMainMenu.Show()
    End Sub

End Class