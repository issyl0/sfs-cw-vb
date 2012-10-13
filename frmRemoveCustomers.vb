Imports System.Data
Imports System.Data.OleDb

Public Class frmRemoveCustomers
    Public accConnection As New OleDbConnection
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        MsgBox("No customers will be removed from the database.")
        Me.Close()
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        'nothing happens yet
        Me.Close()
    End Sub

    Private Sub frmRemoveCustomers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        accConnection = frmMainMenu.accConnection
        Dim strSQL As String = "SELECT cust_surname, cust_forename FROM Customer"
        Dim da As New OleDbDataAdapter(strSQL, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Customer")

        Dim dt As DataTable = ds.Tables(0)
        Dim dr As DataRow

        For Each dr In dt.Rows()
            txtcbCustomerRemove.Items.Add(dr("cust_surname") + ", " + dr("cust_forename"))
        Next

        txtcbCustomerRemove.SelectedIndex = -1
    End Sub

End Class