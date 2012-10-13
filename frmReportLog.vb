Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word

Public Class frmReportLog

    Private Sub frmReportLog_Load(ByVal sender As System.Object, _
                                  ByVal e As System.EventArgs) Handles MyBase.Load

        'Drop down menu populated with months up to 12/12.
        txtcbReportLog.Items.Add("December 2011")
        txtcbReportLog.Items.Add("January 2012")
        txtcbReportLog.Items.Add("February 2012")
        txtcbReportLog.Items.Add("March 2012")
        txtcbReportLog.Items.Add("April 2012")
        txtcbReportLog.Items.Add("May 2012")
        txtcbReportLog.Items.Add("June 2012")
        txtcbReportLog.Items.Add("July 2012")
        txtcbReportLog.Items.Add("August 2012")
        txtcbReportLog.Items.Add("September 2012")
        txtcbReportLog.Items.Add("October 2012")
        txtcbReportLog.Items.Add("November 2012")
        txtcbReportLog.Items.Add("December 2012")

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, _
                            ByVal e As System.EventArgs) Handles btnOK.Click
        'On click, open the selected file from the drop down list with
        'Microsoft Word.

        Dim ms_word As New Word.Application
        Dim worddoc As New Word.Document

        'Make the name shorter purely for ease of typing, despite Visual
        'Studio's tab completion.
        Dim rlsi As String = txtcbReportLog.SelectedItem
        'Make worddoc variable shorter purely for ease of typing in these
        'stressed times where typing every month and every file to open
        'gets tedious and frustrating, but I can't think of a better way
        'to add the months.  Again, tedious and stressful typing 'worddoc'
        'despite Visual Studio's tab completion.
        Dim wd As Word.Document = worddoc

        If rlsi = "December 2011" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_dec11.docx")
        ElseIf rlsi = "January 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_jan12.docx")
        ElseIf rlsi = "February 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_feb12.docx")
        ElseIf rlsi = "March 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_mar12.docx")
        ElseIf rlsi = "April 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_apr12.docx")
        ElseIf rlsi = "May 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_may12.docx")
        ElseIf rlsi = "June 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_jun12.docx")
        ElseIf rlsi = "July 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_jul12.docx")
        ElseIf rlsi = "August 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_aug12.docx")
        ElseIf rlsi = "September 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_sept12.docx")
        ElseIf rlsi = "October 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_oct12.docx")
        ElseIf rlsi = "November 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_nov12.docx")
        ElseIf rlsi = "December 2012" Then
            wd = ms_word.Documents.Open("E:\sfsstuff\log_dec12.docx")
        Else
            MsgBox("No form was selected.")
        End If

        'Let the user see the document when it opens.
        ms_word.WindowState = Word.WdWindowState.wdWindowStateNormal
        ms_word.Visible = True

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, _
                                ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        frmReportMenu.Show()
    End Sub
End Class