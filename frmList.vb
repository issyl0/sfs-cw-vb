Imports System.Data
Imports System.Data.OleDb

Public Class frmList
    Public accConnection As New OleDbConnection

    'I think this is a better way of doing the listing [1] - fewer buttons
    'just a dropdown list which can be added to without having to create
    'another form

    '[1] - compared to the previous separate frmList[Customer|Supplier] etc.

    Private Sub frmList_Load(ByVal sender As System.Object, _
                             ByVal e As System.EventArgs) Handles MyBase.Load

        'Check if db connection is breathing.
        'If it isn't, resuscitate it.  :)
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If

        'Populate the options list box with the following values.
        cbtxtListOptions.Items.Add("Customers")
        cbtxtListOptions.Items.Add("Suppliers")
        cbtxtListOptions.Items.Add("Components")
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) Handles btnShow.Click

        Dim cmdString As String = "" 'Initially blank.
        accConnection = frmLoginForm.accConnection
        Dim ds As New DataSet
        Dim dr As DataRow

        'da0, da1, da2
        'dt0, dt1, dt2
        'Declared within this if statement to make the program recognise 
        'the change in cmdString and so know where to look.
        'Named da<num> and dt<num> to differentiate between different ifs.

        If cbtxtListOptions.Text = "Customers" Then
            cmdString = "SELECT cust_name FROM Customer"
            'Now clear the box if the selection changes.
            lbList.Items.Clear()
            Dim da0 As New OleDbDataAdapter(cmdString, accConnection)
            da0.Fill(ds, "Customer")
            Dim dt0 As DataTable = ds.Tables(0)
            For Each dr In dt0.Rows()
                lbList.Items.Add(dr("cust_name"))
            Next
        ElseIf cbtxtListOptions.Text = "Suppliers" Then
            cmdString = "SELECT supp_name FROM Supplier"
            'Now clear the box if the selection changes.
            lbList.Items.Clear()
            Dim da1 As New OleDbDataAdapter(cmdString, accConnection)
            da1.Fill(ds, "Supplier")
            Dim dt1 As DataTable = ds.Tables(0)
            For Each dr In dt1.Rows()
                lbList.Items.Add(dr("supp_name"))
            Next
        ElseIf cbtxtListOptions.Text = "Components" Then
            cmdString = "SELECT comp_name FROM Component"
            'Now clear the box if the selection changes.
            lbList.Items.Clear()
            Dim da2 As New OleDbDataAdapter(cmdString, accConnection)
            da2.Fill(ds, "Component")
            Dim dt2 As DataTable = ds.Tables(0)
            For Each dr In dt2.Rows()
                lbList.Items.Add(dr("comp_name"))
            Next
        End If

    End Sub

    Private Sub btnShowAssoc_Click(ByVal sender As System.Object, _
                                   ByVal e As System.EventArgs) Handles btnShowAssoc.Click

        If cbtxtListOptions.SelectedItem = "Customers" Then
            MsgBox("No associations to show at this stage.")
        ElseIf cbtxtListOptions.SelectedItem = "Suppliers" Then
            MsgBox("No associations to show at this stage.")
        ElseIf cbtxtListOptions.SelectedItem = "Components" Then
            'Show the association between component and supplier: who supplies
            'each component.
            Call AssocCompSupp()
        End If
    End Sub

    Private Sub AssocCompSupp()

        'Variables here:
        'sn = supplier name;
        'cn = component name.

        accConnection = frmLoginForm.accConnection
        Dim cn As String = lbList.SelectedItem
        Dim compsuppliervalcmd As String = "SELECT comp_supplier FROM Component " & _
                                            "WHERE comp_name = '" & cn & "'"
        Dim sn As String = ""

        Dim da As New OleDbDataAdapter(compsuppliervalcmd, accConnection)
        Dim ds As New DataSet
        da.Fill(ds, "Component")

        Dim compsupplierval As Integer = 0 'Initially.

        'Don't let the program crash if it detects that nothing has been
        'selected - just display the main form anyway.
        Try
            compsupplierval = ds.Tables(0).Rows(0).Item("comp_supplier")
        Catch
        End Try

        'Put the supplier ID value into another variable for ease of
        'use within the function just below.
        Dim csuppv As Integer = compsupplierval

        'Call the function to switch the supplier ID associated with the
        'component in the component table into the actual supplier name.
        'Send the suppname variable SN to the function so that it is
        'easier to reference here afterwards.
        Call SwitchSuppIDToName(csuppv, sn)

        'Make sure that totally blank "is supplied by" don't display
        'blank at both ends because the user hasn't selected an item.
        'This does not cover lack of data about components' suppliers?
        If lbList.SelectedIndex = -1 Then
            MsgBox("Cannot show associations; please select an item from the list.")
        Else
            MsgBox(cn & " is supplied by " & sn & ".")
        End If

    End Sub

    Public Function SwitchSuppIDToName(ByVal csuppval As Integer, _
                                       ByRef finalsn As String)

        accConnection = frmLoginForm.accConnection
        Dim cmdString As String = "SELECT supp_name FROM Supplier WHERE supp_id = " & csuppval & ""

        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet
        da.Fill(ds, "Supplier")

        'Don't let the program crash if it detects that nothing has been
        'selected - just display the main form anyway.
        Try
            finalsn = ds.Tables(0).Rows(0).Item("supp_name")
        Catch
        End Try

        'Functions always return a value, so make this one return one.
        SwitchSuppIDToName = finalsn

    End Function

    Private Sub btnListFullDetails_Click(ByVal sender As System.Object, _
                                         ByVal e As System.EventArgs) Handles btnListFullDetails.Click
        frmListFullDetails.Show()
        Me.Close()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, _
                               ByVal e As System.EventArgs) Handles btnClose.Click
        'Close this window and show the main menu.
        Me.Close()
        frmMainMenu.Show()
    End Sub
End Class