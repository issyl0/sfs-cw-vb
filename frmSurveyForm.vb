Imports System.Data
Imports System.Data.OleDb
Public Class frmSurveyForm

    Public accConnection As New OleDbConnection

    Private Sub frmSurveyForm_Load(ByVal sender As System.Object, _
                                   ByVal e As System.EventArgs) Handles MyBase.Load
        'All this does, for now, is display the selected customer details.
        'In the future, full survey details will be able to be entered by
        'the user, however at this moment I cannot implement this due to 
        'time constraints, so it will be implemented in after-development.
        'To show this to the user when he or she hunts around, there
        'exists a placeholder text label that states that features are
        'coming soon.

        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        accConnection = frmLoginForm.accConnection
        'Input customer name from the other form and attribute it to CN.
        Dim cn As String = frmReportSurvey.txtcbReportSurvey.SelectedItem
        Dim cmdString As String = "SELECT * FROM Customer WHERE cust_name = '" & cn & "'"
        Dim accCommand As New OleDbCommand
        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Customer")

        Dim dt As DataTable = ds.Tables(0)
        Dim dr As DataRow '<- This is redundant.sss

        'Populate the textboxes in the form with existing customer details.
        'All the customer data text boxes are read-only to avoid data error:
        'to edit those, edit the customer details directly, then come back.

        'Populate the customer name textbox with the customer name
        'from the other form, before getting info from db about others.
        txtscn.Text = cn
        If ds.Tables(0).Rows.Count <> 0 Then
            cbsct.Text = ds.Tables(0).Rows(0).Item("cust_title")
            txtscba.Text = ds.Tables(0).Rows(0).Item("cust_billaddress")
            txtscbp.Text = ds.Tables(0).Rows(0).Item("cust_billpostcode")
            txtscia.Text = ds.Tables(0).Rows(0).Item("cust_instaddress")
            txtscip.Text = ds.Tables(0).Rows(0).Item("cust_instpostcode")
            txtschtn.Text = ds.Tables(0).Rows(0).Item("cust_hometelno")
            txtscmtn.Text = ds.Tables(0).Rows(0).Item("cust_mobtelno")
            txtscmpan.Text = ds.Tables(0).Rows(0).Item("cust_mpan")
            txtscea.Text = ds.Tables(0).Rows(0).Item("cust_email")
        End If

        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, _
                               ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
        frmReportSurvey.Show()
    End Sub
End Class