Imports System.Data
Imports System.Data.OleDb

Public Class frmRemoveCustomer
    Public accConnection As New OleDbConnection

    Private Sub frmRemoveCustomers_Load(ByVal sender As System.Object, _
                                        ByVal e As System.EventArgs) Handles MyBase.Load
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        accConnection = frmLoginForm.accConnection
        Dim strSQL As String = "SELECT cust_name FROM Customer"
        Dim da As New OleDbDataAdapter(strSQL, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Customer")

        Dim dt As DataTable = ds.Tables(0)
        Dim dr As DataRow

        For Each dr In dt.Rows()
            txtcbCustomerRemove.Items.Add(dr("cust_name"))
        Next

        txtcbCustomerRemove.SelectedIndex = -1
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnRemove.Click

        Dim cmdString As String = "DELETE * FROM Customer WHERE cust_name = '" & _
                                    Me.txtcbCustomerRemove.SelectedItem & "'"
        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet
        Dim accCommand As New OleDbCommand

        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString

        Call RemoveParameters(accCommand)
        Dim intRemove As Integer

        intRemove = accCommand.ExecuteNonQuery()

        If intRemove = 0 Then
            MsgBox("Sorry, data deletion failed.")
            'Else, assume it went through OK.
        Else
            btnRemove.Enabled = False
        End If

    End Sub

    Private Sub RemoveParameters(ByRef acccmd As OleDbCommand)
        acccmd.Parameters.Add("@cust_name", OleDbType.Char).Value = _
                                Me.txtcbCustomerRemove.SelectedItem
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        frmRemove.Show()
    End Sub

End Class