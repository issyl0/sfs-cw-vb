Imports System.Data
Imports System.Data.OleDb

Public Class frmListFullDetails
    Public accConnection As New OleDbConnection

    Private Sub frmListFullDetails_Load(ByVal sender As System.Object, _
                                        ByVal e As System.EventArgs) Handles MyBase.Load

        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        If frmList.cbtxtListOptions.SelectedItem = "Customers" Then
            Call PopulateCustInfo()
        ElseIf frmList.cbtxtListOptions.SelectedItem = "Suppliers" Then
            Call PopulateSuppInfo()
        ElseIf frmList.cbtxtListOptions.SelectedItem = "Components" Then
            Call PopulateCompInfo()
        End If
    End Sub

    Private Sub PopulateCustInfo()

        accConnection = frmLoginForm.accConnection
        'Input customer name from the other form and attribute it to CTN.
        Dim ctn As String = frmList.lbList.SelectedItem 'Customer name.
        Dim cmdString As String = "SELECT * FROM Customer WHERE cust_name = '" & ctn & "'"
        Dim accCommand As New OleDbCommand
        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Customer")

        Dim dt As DataTable = ds.Tables(0)
        'Dim dr As DataRow <- This is redundant - it does not do anything.

        'Populate the textboxes in the form with existing customer details.
        'All the customer data text boxes are read-only to avoid data error:
        'to edit those, edit the customer details directly, then come back.

        'Populate the customer name textbox with the customer name
        'from the other form, before getting info from db about others.

        'An example of my variable naming here:
        'The variable 'txtldcn' = textbox listdetails custname.
        txtldctn.Text = ctn
        If ds.Tables(0).Rows.Count <> 0 Then
            cbldctt.Text = ds.Tables(0).Rows(0).Item("cust_title")
            txtldctba.Text = ds.Tables(0).Rows(0).Item("cust_billaddress")
            txtldctbp.Text = ds.Tables(0).Rows(0).Item("cust_billpostcode")
            txtldctia.Text = ds.Tables(0).Rows(0).Item("cust_instaddress")
            txtldctip.Text = ds.Tables(0).Rows(0).Item("cust_instpostcode")
            txtldcthtn.Text = ds.Tables(0).Rows(0).Item("cust_hometelno")
            txtldctmtn.Text = ds.Tables(0).Rows(0).Item("cust_mobtelno")
            txtldctea.Text = ds.Tables(0).Rows(0).Item("cust_email")
            txtldctmpan.Text = ds.Tables(0).Rows(0).Item("cust_mpan")
        End If

        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString
    End Sub

    Private Sub PopulateSuppInfo()

        accConnection = frmLoginForm.accConnection
        'Input supplier name from the other form and attribute it to SPN.
        Dim spn As String = frmList.lbList.SelectedItem 'Supplier name.
        Dim cmdString As String = "SELECT * FROM Supplier WHERE supp_name = '" & spn & "'"
        Dim accCommand As New OleDbCommand
        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Supplier")

        Dim dt As DataTable = ds.Tables(0)
        'Dim dr As DataRow <- This is redundant - it does not do anything.

        'Populate the textboxes in the form with existing supplier details.
        'All the supplier data text boxes are read-only to avoid data error:
        'to edit those, edit the supplier details directly, then come back.

        'Populate the supplier name textbox with the customer name
        'from the other form, before getting info from db about others.

        'An example of my variable naming here:
        'The variable 'txtldspn' = textbox listdetails suppname.
        txtldspn.Text = spn
        If ds.Tables(0).Rows.Count <> 0 Then
            txtldspa.Text = ds.Tables(0).Rows(0).Item("supp_address")
            txtldspp.Text = ds.Tables(0).Rows(0).Item("supp_postcode")
            txtldsptn.Text = ds.Tables(0).Rows(0).Item("supp_telno")
            txtldspc.Text = ds.Tables(0).Rows(0).Item("supp_contactname")
        End If

        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString

    End Sub

    Private Sub PopulateCompInfo()

        accConnection = frmLoginForm.accConnection
        'Input component name from the other form and attribute it to CPN.
        Dim cpn As String = frmList.lbList.SelectedItem 'Component name.
        Dim cmdString As String = "SELECT * FROM Component WHERE comp_name = '" & cpn & "'"

        Dim accCommand As New OleDbCommand
        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Component")

        Dim dt As DataTable = ds.Tables(0)
        'Dim dr As DataRow <- This is redundant - it does not do anything.

        Dim compsupplierval As Integer = 0

        'Check if a supplier is selected.  Before this Try/Catch, the
        'program would crash due to no value at row 0 of the table.
        Try
            compsupplierval = ds.Tables(0).Rows(0).Item("comp_supplier")
        Catch
        End Try
        Dim csuppv As Integer = compsupplierval
        Dim sn As String = "" 'Supplier name, initially blank here.

        'Now call the List form's switching supplier ID function.
        'No point creating a near identical one if there exists the
        'ability of reusing code.
        Call frmList.SwitchSuppIDToName(csuppv, sn)

        'Populate the textboxes in the form with existing component details.
        'All the component data text boxes are read-only to avoid data error:
        'to edit those, edit the component details directly, then come back.

        'Populate the component name textbox with the customer name
        'from the other form, before getting info from db about others.

        'An example of my variable naming here:
        'The variable 'txtldcpn' = textbox listdetails compname.
        txtldcpn.Text = cpn
        If ds.Tables(0).Rows.Count <> 0 Then
            txtldcpt.Text = ds.Tables(0).Rows(0).Item("comp_type")
            txtldcpsn.Text = ds.Tables(0).Rows(0).Item("comp_serialno")
            txtldcppkwp.Text = ds.Tables(0).Rows(0).Item("comp_panelkwp")
            txtldcpspn.Text = sn
        End If

        accCommand.Connection = frmLoginForm.accConnection
        accCommand.CommandType = CommandType.Text
        accCommand.CommandText = cmdString
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, _
                               ByVal e As System.EventArgs) Handles btnClose.Click
        frmList.Show()
        Me.Close()
    End Sub

End Class