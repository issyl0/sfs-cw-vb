Imports System.Data
Imports System.Data.OleDb

Class frmRemoveComponent
    Public accConnection As New OleDbConnection

    Private Sub frmRemoveComponent_Load(ByVal sender As System.Object, _
                                        ByVal e As System.EventArgs) Handles MyBase.Load
        'Check if the database connection is breathing.
        'If it isn't, resuscitate it.  :-)
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        accConnection = frmLoginForm.accConnection
        Dim strSQL As String = "SELECT comp_name FROM Component"
        Dim da As New OleDbDataAdapter(strSQL, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Component")

        Dim dt As DataTable = ds.Tables(0)
        Dim dr As DataRow

        For Each dr In dt.Rows()
            'List component names in the box.
            txtcbComponentRemove.Items.Add(dr("comp_name"))
        Next

        txtcbComponentRemove.SelectedIndex = -1

    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnRemove.Click

        Dim cmdString As String = "DELETE * FROM Component WHERE comp_name = " & _
                                    Me.txtcbComponentRemove.SelectedItem & ""
        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet
        Dim accCommand As New OleDbCommand
        Dim intRemove As Integer

        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString

        intRemove = accCommand.ExecuteNonQuery()

        If intRemove = 0 Then
            MsgBox("Data deletion failed.")
            'Else, assume it went through OK.
        Else
            btnRemove.Enabled = False
        End If

    End Sub

    Private Sub RemoveParameters(ByRef acccmd As OleDbCommand)
        acccmd.Parameters.Add("@comp_name", OleDbType.Char).Value = _
                                txtcbComponentRemove.SelectedItem
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        frmRemove.Show()
    End Sub

End Class