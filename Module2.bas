Attribute VB_Name = "Module2"
Sub CategorizeEmailsByAmount()
    Dim objNamespace As Outlook.NameSpace
    Dim objCategory As Outlook.Category
    Dim olItem As Outlook.MailItem
    Dim emailSubject As String
    Dim invoiceAmount As Double
    Dim filePath As String
    Dim ts As Object
    Dim line As String
    
    filePath = "C:\Users\b.granderath\OneDrive - Wünsche Group\Dokumente\Outlook SQL Scrip\sqlscript.txt"
    If xmlInvoiceAmount = 0 Then
        Set objNamespace = Application.GetNamespace("MAPI")
        Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(filePath, 1)
            For Each olItem In Application.ActiveExplorer.Selection
                If TypeOf olItem Is Outlook.MailItem Then
                    Do While Not ts.AtEndOfStream
                        line = ts.ReadLine
                        line = Trim(line)
                        Debug.Print "Kein XML Invoice"
                        If IsNumeric(line) And line <> "" Then
                            invoiceAmount = CDbl(line)
                            emailSubject = GetEmailSubjectFromInvoiceAmount(invoiceAmount)
                            Debug.Print emailSubject
                            Debug.Print invoiceAmount
                            On Error Resume Next
                            Set objCategory = objNamespace.Categories.Item(emailSubject)
                            On Error GoTo 0
                            
                             If objCategory Is Nothing Then
                                Set objCategory = objNamespace.Categories.Add(emailSubject)
                            End If
                            
                            If InStr(1, olItem.Categories, objCategory.Name) = 0 Then
                                If Len(olItem.Categories) > 0 Then
                                    olItem.Categories = olItem.Categories & ";" & objCategory.Name
                                Else
                                    olItem.Categories = objCategory.Name
                                End If
                            End If
                            olItem.Save
                            Exit Do
                        End If
                    Loop
                End If
            Next olItem
        ts.Close
    Else
        Set objNamespace = Application.GetNamespace("MAPI")
        For Each olItem In Application.ActiveExplorer.Selection
            If TypeOf olItem Is Outlook.MailItem Then
            
                Debug.Print "XML Invoice"
                emailSubject = GetEmailSubjectFromInvoiceAmount(xmlInvoiceAmount)
                Debug.Print emailSubject
                Debug.Print xmlInvoiceAmount
                On Error Resume Next
                
                Set objCategory = objNamespace.Categories.Item(emailSubject)
                On Error GoTo 0
                
                If objCategory Is Nothing Then
                    Set objCategory = objNamespace.Categories.Add(emailSubject)
                End If
                
                If InStr(1, olItem.Categories, objCategory.Name) = 0 Then
                    If Len(olItem.Categories) > 0 Then
                        olItem.Categories = olItem.Categories & ";" & objCategory.Name
                    Else
                        olItem.Categories = objCategory.Name
                    End If
                End If
                
                olItem.Save
            End If
        Next olItem
    End If
End Sub




Function GetEmailSubjectFromInvoiceAmount(ByVal amount As Double) As String
    If amount < 5000 Then
        GetEmailSubjectFromInvoiceAmount = "bis 5 TSD"
    ElseIf amount >= 5000 And amount < 15000 Then
        GetEmailSubjectFromInvoiceAmount = "5-15 TSD"
    ElseIf amount >= 15000 And amount <= 50000 Then
        GetEmailSubjectFromInvoiceAmount = "15-50 TSD"
    Else
        GetEmailSubjectFromInvoiceAmount = "über 50 TSD"
    End If
End Function
