Attribute VB_Name = "Module1"
Public xmlInvoiceAmount As Double
Sub GetRechnungsnummern()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim olSelection As Outlook.Selection
    Dim olItem As Object
    Dim emailBody As String
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim idPattern As String
    Dim idMatches As Object
    Dim sqlScript As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim isFirstMatch As Boolean
    Dim invoiceIDs As String
    Dim orderByClause As String
    Dim matchCounter As Integer
    Dim powershellPath As String
    Dim xmlAttachment As Outlook.Attachment
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim originalAmount As String
    
    Set olApp = New Outlook.Application
    Set olSelection = olApp.ActiveExplorer.Selection
    idPattern = "(AR|DN)\d+"
    filePath = "C:\Users\b.granderath\OneDrive - Wünsche Group\Dokumente\Outlook SQL Scrip\sqlscript.sql"
    powershellPath = "C:\Users\b.granderath\OneDrive - Wünsche Group\Dokumente\Outlook SQL Scrip\sqlscript.ps1"
    isFirstMatch = True
    matchCounter = 1

    fileNumber = FreeFile
    Open filePath For Output As fileNumber
       sqlScript = "SELECT ROUND(CAST([Rechnungsbetrag] AS INT), 1)" & vbCrLf _
                & "FROM [wsmb].[dbo].[LOG_AX_RECHNUNGSERFASSUNG]" & vbCrLf _
                & "WHERE Beleg IN ([placeholder])" & vbCrLf _
                & "ORDER BY CASE Beleg"

    For Each olItem In olSelection
        If TypeOf olItem Is Outlook.MailItem Then
            Set olMail = olItem
            emailBody = olMail.Body
            Set idMatches = GetRegExpMatches(emailBody, idPattern)
            For Each xmlAttachment In olMail.Attachments
                If Right(xmlAttachment.FileName, 3) = "xml" Then
                    Debug.Print xmlAttachment.FileName
                    Close fileNumber
                    xmlAttachment.SaveAsFile "C:\Temp\" & xmlAttachment.FileName
                    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
                    xmlDoc.Async = False
                    xmlDoc.Load "C:\Temp\" & xmlAttachment.FileName
                
                    Set xmlNode = xmlDoc.SelectSingleNode("//InvoiceAmount")
                    If Not xmlNode Is Nothing Then
                        Debug.Print "XML mit InvoiceAmount Node erkannt"
                        originalAmount = xmlNode.Text
                        originalAmount = Replace(originalAmount, ".", ",")
                        xmlInvoiceAmount = CDbl(originalAmount)
                        Debug.Print xmlInvoiceAmount
                    Else
                        MsgBox "Kein InvoiceAmount in der XML gefunden."
                        Exit Sub
                    End If
                    Kill "C:\Temp\" & xmlAttachment.FileName
                    Call Module2.CategorizeEmailsByAmount
                End If
            Next xmlAttachment
            If Not idMatches Is Nothing And idMatches.Count > 0 Then
                Debug.Print "Invoice in Body erkannt."
                For Each match In idMatches
                    If isFirstMatch Then
                        invoiceIDs = "'" & match.Value & "'"
                        orderByClause = orderByClause & vbCrLf & " WHEN '" & match.Value & "' THEN " & matchCounter
                        isFirstMatch = False
                    Else
                        invoiceIDs = invoiceIDs & ", '" & match.Value & "'"
                        orderByClause = orderByClause & vbCrLf & " WHEN '" & match.Value & "' THEN " & matchCounter
                    End If
                    matchCounter = matchCounter + 1
                Next match
            End If
        End If
    Next olItem
    For Each xmlAttachment In olMail.Attachments
        If Right(xmlAttachment.FileName, 3) = "xml" Then
            Exit Sub
        End If
    Next xmlAttachment
    orderByClause = orderByClause & vbCrLf & " ELSE 999 END;"
    sqlScript = Replace(sqlScript, "[placeholder]", invoiceIDs)
    sqlScript = sqlScript & orderByClause
    
    Print #fileNumber, sqlScript

    Close fileNumber
    wsh.Run "powershell.exe -ExecutionPolicy Bypass -File """ & powershellPath & """", vbNormalFocus, waitOnReturn
    Call Module2.CategorizeEmailsByAmount
    
    Kill "C:\Users\b.granderath\OneDrive - Wünsche Group\Dokumente\Outlook SQL Scrip\sqlscript.txt"
    MsgBox olSelection.Count & " Email(s) nach Rechnungsbetrag kategorisiert."
End Sub

Function GetRegExpMatches(ByVal inputString As String, ByVal pattern As String) As Object
    Dim regExp As Object
    Dim matches As Object

    Set regExp = CreateObject("VBScript.RegExp")
    regExp.Global = True
    regExp.IgnoreCase = True
    regExp.pattern = pattern

    Set matches = regExp.Execute(inputString)
    Set GetRegExpMatches = matches
End Function

