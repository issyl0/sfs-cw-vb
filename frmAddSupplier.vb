Imports System.Data
Imports System.Data.OleDb

Public Class frmAddSupplier
    Public accConnection As New OleDbConnection

    Private Sub frmAddSupplier_Load(ByVal sender As System.Object, _
                                    ByVal e As System.EventArgs) Handles MyBase.Load
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnSave.Click

        Dim cmdString As String = "INSERT INTO Supplier (supp_name, supp_address, supp_postcode, " _
                                  & "supp_telno, supp_contactname)" _
                                  & "VALUES (@supp_name,@supp_address,@supp_postcode," _
                                  & "@supp_telno,@supp_contactname)"

        Dim AddSupplierDataAdapter As New OleDbDataAdapter
        Dim accCommand As New OleDbCommand
        Dim intInsert As Integer

        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString
        Call InsertParameters(accCommand)
        'Now check if all the textboxes are populated.
        Try
            intInsert = accCommand.ExecuteNonQuery()
        Catch ex As Exception
            intInsert = 0
            MsgBox("Enter a value in each of the boxes, and make it a valid one!")
        End Try

        Call frmSwankyCode.CheckAdditions(intInsert, btnSave)

    End Sub

    Private Sub InsertParameters(ByRef acccmd As OleDbCommand)
        acccmd.Parameters.Add("@supp_name", OleDbType.Char).Value = txtSupplierName.Text
        acccmd.Parameters.Add("@supp_address", OleDbType.Char).Value = txtSupplierAddress.Text
        acccmd.Parameters.Add("@supp_postcode", OleDbType.Char).Value = txtSupplierPostCode.Text
        acccmd.Parameters.Add("@supp_telno", OleDbType.Char).Value = mtxtSupplierTelNo.Text
        acccmd.Parameters.Add("@supp_contactname", OleDbType.Char).Value = txtSupplierContactName.Text

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnCancel.Click

        If btnSave.Enabled = False Then
            'It's all fine and has gone through OK, so don't display a
            'message because that would just be annoying, just close this
            'and display the previous form.
            Me.Close()
            frmAdd.Show()
        Else
            'If it hasn't gone through OK, or nothing has been added,
            'then let the user know this so as not to panic them;
            'if the user did not mean to click the button, he/she knows.
            MsgBox("This window will close and these details will not be saved.")
            Me.Close()
            frmAdd.Show()
        End If

    End Sub

End Class