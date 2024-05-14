Attribute VB_Name = "Module3"
Public Sub SaveToAutoDispo()
    Dim Explorer As Outlook.Explorer
    Dim Selection As Outlook.Selection
    Dim i As Integer
    Dim MailItem As Outlook.MailItem
    Dim Atmt As Outlook.Attachment
    Dim FileName As String
    Dim FolderPath As String
    
    FolderPath = "\\wuensche-group.local\Applications\AXPROD\EDI\AX_Import\ORDERS\Centershop\"
    
    Set Explorer = Application.ActiveExplorer
    Set Selection = Explorer.Selection
    
    For i = 1 To Selection.Count
        If TypeOf Selection.Item(i) Is Outlook.MailItem Then
            Set MailItem = Selection.Item(i)
            
            For Each Atmt In MailItem.Attachments
                If Right(Atmt.FileName, 4) = ".txt" Then
                    FileName = FolderPath & Atmt.FileName
                    Atmt.SaveAsFile FileName
                End If
            Next Atmt
        End If
    Next i
    MsgBox Selection.Count & " Nachrichten unter " & FolderPath & " gespeichert."
End Sub
