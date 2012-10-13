Imports System.Data
Imports System.Data.OleDb

Public Class frmAddCustomer
    Public accConnection As New OleDbConnection

    Private Sub frmAddCustomers_Load(ByVal sender As System.Object, _
                                     ByVal e As System.EventArgs) Handles MyBase.Load
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        'Populate this list with selected titles.
        Me.cbCustomerTitle.Items.Add("Mr")
        Me.cbCustomerTitle.Items.Add("Mrs")
        Me.cbCustomerTitle.Items.Add("Miss")
        Me.cbCustomerTitle.Items.Add("Dr")

        'Add tooltip to lblCustomerName and txtCustomerName field 
        'so that the user knows to add the name in a uniform way.
        Me.ttCustomerName.SetToolTip(Me.lblCustomerName, _
                                     "Name must be in the format 'Forename <space> Surname'.")
        Me.ttCustomerName.SetToolTip(Me.txtCustomerName, _
                                     "Name must be in the format 'Forename <space> Surname'.")
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnSave.Click

        Dim cmdString As String = "INSERT INTO Customer (cust_title, cust_name, " _
                                  & "cust_billaddress, cust_billpostcode, cust_instaddress, " _
                                  & "cust_instpostcode, cust_hometelno, cust_mobtelno, " _
                                  & "cust_email, cust_mpan) VALUES (@cust_title,@cust_name," _
                                  & "@cust_billaddress,@cust_billpostcode,@cust_instaddress," _
                                  & "@cust_instpostcode,@cust_hometelno,@cust_mobtelno," _
                                  & "@cust_email,@cust_mpan)"

        Dim AddCustomerDataAdapter As New OleDbDataAdapter
        Dim accCommand As New OleDbCommand
        Dim intInsert As Integer
        Dim CustomerForename As String

        CustomerForename = txtCustomerName.Text
        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString

        'Make the customer bill address be the install address too if
        'the user ticks the box.
        If cbCustomerBillInstallSameAddress.CheckState = 1 Then
            txtCustomerInstAddress.Text = txtCustomerBillAddress.Text
            txtCustomerInstPostcode.Text = txtCustomerBillPostcode.Text
        End If

        Call InsertParameters(accCommand)

        'Check if there are no empty textboxes.  If there aren't, do all
        'other checks, and send it through to the database.
        Try
            intInsert = accCommand.ExecuteNonQuery()
        Catch ex As Exception
            intInsert = 0
            MsgBox("Enter a value in each of the boxes, and make it a valid one!")
        End Try

        'Check if the email textbox contains '@', therefore if it is a
        'valid email address.
        Dim pos As Integer
        Dim etb As String = txtCustomerEmail.Text 'email textbox
        Dim atsymbol As String = "@"
        pos = InStr(etb, atsymbol)
        If pos = 0 Then
            MsgBox("Input another email address!")
            intInsert = 0
        End If

        Call frmSwankyCode.CheckAdditions(intInsert, btnSave)

    End Sub

    Private Sub InsertParameters(ByRef acccmd As OleDbCommand)
        acccmd.Parameters.Add("@cust_title", OleDbType.Char).Value = _
                                cbCustomerTitle.SelectedItem
        acccmd.Parameters.Add("@cust_name", OleDbType.Char).Value = _
                                txtCustomerName.Text
        acccmd.Parameters.Add("@cust_billaddress", OleDbType.Char).Value = _
                                txtCustomerBillAddress.Text
        acccmd.Parameters.Add("@cust_billpostcode", OleDbType.Char).Value = _
                                txtCustomerBillPostcode.Text
        acccmd.Parameters.Add("@cust_instaddress", OleDbType.Char).Value = _
                                txtCustomerInstAddress.Text
        acccmd.Parameters.Add("@cust_instpostcode", OleDbType.Char).Value = _
                                txtCustomerInstPostcode.Text
        acccmd.Parameters.Add("@cust_hometelno", OleDbType.Char).Value = _
                                txtCustomerHomeTelNo.Text
        acccmd.Parameters.Add("@cust_mobtelno", OleDbType.Char).Value = _
                                mtxtCustomerMobTelNo.Text
        acccmd.Parameters.Add("@cust_email", OleDbType.Char).Value = _
                                txtCustomerEmail.Text
        acccmd.Parameters.Add("@cust_mpan", OleDbType.Char).Value = _
                                txtCustomerMpanNo.Text
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