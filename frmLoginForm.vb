Imports System.Data
Imports System.Data.OleDb

Public Class frmLoginForm
    Public accConnection As New OleDbConnection
    Dim ConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source = E:\SFS.mdb"

    Private Sub frmLoginForm_Load(ByVal sender As System.Object, _
                                  ByVal e As System.EventArgs) Handles MyBase.Load
        accConnection = New OleDbConnection(ConnectionString)
    End Sub

    Private Sub OK_Click(ByVal sender As System.Object, _
                         ByVal e As System.EventArgs) Handles OK.Click

        Dim cmdString As String = "SELECT login_uname, login_pword FROM Login"
        Dim da As New OleDbDataAdapter(cmdString, accConnection)
        Dim ds As New DataSet
        Dim u As String = txtUsername.Text
        Dim p As String = txtPassword.Text

        da.Fill(ds, "Login")
        Dim dt As DataTable = ds.Tables(0)

        'The usernames and passwords obtained from the database, to
        'compare with the ones the user entered.
        Dim loginuname As String = ""
        Dim loginpword As String = ""

        If ds.Tables(0).Rows.Count <> 0 Then
            loginuname = ds.Tables(0).Rows(0).Item("login_uname")
            loginpword = ds.Tables(0).Rows(0).Item("login_pword")
        End If

        If u = loginuname And p = loginpword Then
            frmMainMenu.Show()
        Else
            MsgBox("Incorrect username or password.  Try again.")
            'Error, then blank the textboxes so that the user doesn't
            'have to.
            txtUsername.Text = ""
            txtPassword.Text = ""
        End If

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, _
                             ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
