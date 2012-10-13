Imports System.Data
Imports System.Data.OleDb

Class frmReportInvoice
    Public accConnection As New OleDbConnection

    Private Sub frmReportInvoice_Load(ByVal sender As System.Object, _
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
            'List customer names in the box so that the user can select
            'a customer to show the invoice of.  At least customer details
            'will be inserted into the Word document at this stage.
            txtcbCustNameForInvoice.Items.Add(dr("cust_name"))
        Next

        txtcbCustNameForInvoice.SelectedIndex = -1

    End Sub


    Private Sub btnOK2_Click(ByVal sender As System.Object, _
                             ByVal e As System.EventArgs) Handles btnOK2.Click

        accConnection = frmLoginForm.accConnection
        'Input supplier name from the other form and attribute it to CTN.
        Dim ctn As String = txtcbCustNameForInvoice.SelectedItem
        Dim cmdString As String = "SELECT * FROM Customer WHERE cust_name = '" & ctn & "'"

        Dim accCommand As New OleDbCommand
        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Customer")

        Dim dt As DataTable = ds.Tables(0)

        'Now create variables to hold the useful address values of the
        ''* FROM Customer' so they'll be useful later on for the Word
        'invoicing stuff.
        'aicn, for example, = add invoice cust name
        Dim aict As String = ds.Tables(0).Rows(0).Item("cust_title")
        Dim aicn As String = ds.Tables(0).Rows(0).Item("cust_name")
        'Billing address and postcode in this case.
        Dim aica As String = ds.Tables(0).Rows(0).Item("cust_billaddress")
        Dim aicp As String = ds.Tables(0).Rows(0).Item("cust_billpostcode")

        'Now for some Word automation magic.  Call a procedure to do this
        'to save cluttering this OK button's execution with code.
        Call frmSwankyCode.WordInvoiceAutomationMagic(aict, aicn, aica, aicp)

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnCancel.Click
        frmReportMenu.Show()
        Me.Close()
    End Sub
End Class