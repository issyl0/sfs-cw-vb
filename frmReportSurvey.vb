Imports System.Data
Imports System.Data.OleDb

Public Class frmReportSurvey

    Private Sub frmReportSurvey_Load(ByVal sender As System.Object, _
                                     ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cmdString As String = "SELECT cust_name FROM Customer"
        'check if db connection is breathing
        'if it isn't, resuscitate it :)
        If frmLoginForm.accConnection.State <> ConnectionState.Open Then
            frmLoginForm.accConnection.Open()
        End If
        Dim accConnection As OleDbConnection = frmLoginForm.accConnection
        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet

        da.Fill(ds, "Customer")

        Dim dt As DataTable = ds.Tables(0)
        Dim dr As DataRow

        For Each dr In dt.Rows()
            txtcbReportSurvey.Items.Add(dr("cust_name"))
        Next

        txtcbReportSurvey.SelectedIndex = -1

    End Sub


    Private Sub btnOK_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs) Handles btnOK.Click

        Dim custname As String = txtcbReportSurvey.SelectedItem
        frmSurveyForm.Show()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        frmReportMenu.Show()
    End Sub
End Class