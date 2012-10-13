Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word

Public Class frmSwankyCode

    'This is a form because I couldn't work out how to make it work any
    'other way.

    'This is where all the code goes that is referenced from other subs
    'in other forms, to save space and attempt to minimise repetition.

    Public Sub CheckAdditions(ByVal intInsert As Integer, _
                              ByRef btnSave As Button)

        If intInsert = 0 Then
            MsgBox("Data insertion failed.")
        Else
            'Assume it went through OK and disable the Save button to
            'guard against duplication.
            btnSave.Enabled = False
        End If

    End Sub

    Public Sub WordInvoiceAutomationMagic(ByVal ict As String, ByVal icn As String, _
                                   ByVal ica As String, ByVal icp As String)

        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim CustAddressStuff As Word.Paragraph
        Dim CustInvoiceHeader As Word.Paragraph

        'Start Word and open the document.
        'This is object stuff that I don't understand, but it was found on
        'a Microsoft doc, and the other stuff I tried didn't work, so I
        'used it.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert the customer billing address stuff.
        CustAddressStuff = oDoc.Content.Paragraphs.Add
        CustAddressStuff.Range.Text = ict & " " & icn & Chr(11) & ica _
                                      & Chr(11) & icp & Chr(11)
        CustAddressStuff.Range.ParagraphFormat.Alignment = _
            Word.WdParagraphAlignment.wdAlignParagraphRight
        CustAddressStuff.Range.Font.Bold = False

        'Sort out the spacing afterwards.
        CustAddressStuff.Format.SpaceAfter = 4
        CustAddressStuff.Range.InsertParagraphAfter()

        'Now, after the spacing, the INVOICE header, centred.
        CustInvoiceHeader = _
            oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        CustInvoiceHeader.Range.Text = "INVOICE"
        CustInvoiceHeader.Range.ParagraphFormat.Alignment = _
            Word.WdParagraphAlignment.wdAlignParagraphCenter
        CustInvoiceHeader.Range.Font.Size = 22
        CustInvoiceHeader.Range.Font.Bold = True

    End Sub

    Public Sub WordQuotationAutomationMagic(ByVal qct As String, ByVal qcn As String, _
                                   ByVal qca As String, ByVal qcp As String)

        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim CustAddressStuff As Word.Paragraph
        Dim CustQuotationHeader As Word.Paragraph

        'Start Word and open the document.
        'This is object stuff that I don't understand, but it was found on
        'a Microsoft doc, and the other stuff I tried didn't work, so I
        'used it.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert the customer billing address stuff.
        CustAddressStuff = oDoc.Content.Paragraphs.Add
        CustAddressStuff.Range.Text = qct & " " & qcn & Chr(11) & qca _
                                      & Chr(11) & qcp & Chr(11)
        CustAddressStuff.Range.ParagraphFormat.Alignment = _
            Word.WdParagraphAlignment.wdAlignParagraphRight
        CustAddressStuff.Range.Font.Bold = False

        'Sort out the spacing afterwards.
        CustAddressStuff.Format.SpaceAfter = 4
        CustAddressStuff.Range.InsertParagraphAfter()

        'Now, after the spacing, the Quotation header, centred.
        CustQuotationHeader = _
            oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        CustQuotationHeader.Range.Text = "QUOTATION"
        CustQuotationHeader.Range.ParagraphFormat.Alignment = _
            Word.WdParagraphAlignment.wdAlignParagraphCenter
        CustQuotationHeader.Range.Font.Size = 22
        CustQuotationHeader.Range.Font.Bold = True

    End Sub

End Class