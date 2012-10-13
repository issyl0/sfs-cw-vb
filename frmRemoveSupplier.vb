Imports System.Data
Imports System.Data.OleDb

Public Class frmRemoveSupplier
    Public accConnection As New OleDbConnection

    Private Sub frmRemoveSupplier_Load(ByVal sender As System.Object, _
                                       ByVal e As System.EventArgs) Handles MyBase.Load

        'Check if db connection is breathing.
        'If it isn't, resuscitate it.
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        accConnection = frmLoginForm.accConnection
        Dim strSQL As String = "SELECT supp_id,supp_name FROM Supplier"
        Dim da As New OleDbDataAdapter(strSQL, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Supplier")

        Dim dt As DataTable = ds.Tables(0)
        Dim dr As DataRow

        For Each dr In dt.Rows()
            'Populate the box with the supplier names.
            txtcbSupplierRemove.Items.Add(dr("supp_name"))
        Next

        txtcbSupplierRemove.SelectedIndex = -1

    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnRemove.Click

        Dim RemoveSupplierDataAdapter As New OleDbDataAdapter
        Dim accCommand As New OleDbCommand
        Dim cmdString As String = "DELETE * FROM Supplier WHERE supp_name = " & _
                                    Me.txtcbSupplierRemove.SelectedItem & ""

        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString

        Call RemoveParameters(accCommand)
        'Execute the query to delete the record from the db.
        Dim intRemove As Integer

        intRemove = accCommand.ExecuteNonQuery()

        If intRemove = 0 Then
            MsgBox("Data deletion failed.")
            'Else, assume it went through OK.
        Else
            btnRemove.Enabled = False
        End If

    End Sub

    Private Sub RemoveParameters(ByRef acccmd As OleDbCommand)
        acccmd.Parameters.Add("@supp_name", OleDbType.Char).Value = _
                                    txtcbSupplierRemove.SelectedItem
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        frmRemove.Show()
    End Sub
End Class