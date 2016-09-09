Public Sub ReportTool(itm As Outlook.MailItem)
    Dim objAtt As Outlook.Attachment
    Dim saveFolder As String
    Dim dateFormat
    dateFormat = Format(itm.ReceivedTime, "YYYY-MM-DD")
    saveFolder = "S:\Rotinas - Relatórios\Teste Vendas Diretas\Acompanhamento de Credenciamentos"
    For Each objAtt In itm.Attachments
        objAtt.SaveAsFile saveFolder & "\" & dateFormat & " " & objAtt.DisplayName
    Next
End Sub
