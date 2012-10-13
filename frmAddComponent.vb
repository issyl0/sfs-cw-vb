Imports System.Data
Imports System.Data.OleDb

Class frmAddComponent
    Public accConnection As New OleDbConnection

    Private Sub AddComponent_Load(ByVal sender As System.Object, _
                                  ByVal e As System.EventArgs) Handles MyBase.Load
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        accConnection = frmLoginForm.accConnection
        Dim strSQL As String = "SELECT supp_name FROM Supplier"
        Dim da As New OleDbDataAdapter(strSQL, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Supplier")

        Dim dt As DataTable = ds.Tables(0)
        Dim dr As DataRow

        For Each dr In dt.Rows()
            'List supplier names in the box so that the user can select
            'the supplier that the component is coming from.  Pull this
            'from the Supplier table.
            txtcbCompSupplierName.Items.Add(dr("supp_name"))
        Next

        txtcbCompSupplierName.SelectedIndex = -1

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnSave.Click

        'This takes the supplier name from the supplier table and gets the
        'supplier ID from it, thereby letting it into the Component table.

        Dim suppidvaluecmd As String = "SELECT supp_id FROM Supplier WHERE " _
                                       & "supp_name = '" & txtcbCompSupplierName.SelectedItem & "'"

        Dim cmdString As String = "INSERT INTO Component (comp_name, comp_type, comp_serialno," _
                                  & " comp_panelkwp, comp_supplier)" _
                                  & "VALUES (@comp_name,@comp_type,@comp_serialno," _
                                  & "@comp_panelkwp,@comp_supplier)"

        Dim da As New OleDbDataAdapter(suppidvaluecmd, accConnection)
        Dim ds As New DataSet
        da.Fill(ds, "Supplier")

        Dim AddSupplierDataAdapter As New OleDbDataAdapter
        Dim accCommand As New OleDbCommand
        Dim intInsert As Integer

        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString
        Dim supplieridvalue As Integer = ds.Tables(0).Rows(0).Item("supp_id")
        MsgBox(supplieridvalue)
        Call InsertParameters(accCommand, supplieridvalue)
        'Now check if all the textboxes are populated.
        Try
            intInsert = accCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Enter a value in each of the boxes, and make it a valid one!")
        End Try

        Call frmSwankyCode.CheckAdditions(intInsert, btnSave)
    End Sub

    Private Sub InsertParameters(ByRef acccmd As OleDbCommand, ByRef siv As Integer)

        acccmd.Parameters.Add("@comp_name", OleDbType.Char).Value = txtComponentName.Text
        acccmd.Parameters.Add("@comp_type", OleDbType.Char).Value = txtComponentType.Text
        acccmd.Parameters.Add("@comp_serialno", OleDbType.Char).Value = txtComponentSerialNo.Text
        acccmd.Parameters.Add("@comp_panelkwp", OleDbType.Numeric).Value = txtComponentPanelkWp.Text
        'And now the supplier ID value.
        acccmd.Parameters.Add("@comp_supplier", OleDbType.Numeric).Value = siv

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