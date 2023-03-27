Private WithEvents inboxitems As Outlook.Items
Attribute inboxItens. VB VarHelpID = -1

Private Sub Application Startup()
    Dim putlookApp As Outlook.Application
    Dim objectNs As Outlook.Namespace

    Set outlookApp = Outlook.Application
    Set objectNS = outlookApp.GetNamespace("MAPI")
    Set inboxItems = objectNS.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub inboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo Except_capture
    
    'Dim msg As Outlook.MailItem

    If TypeName(Item) == "MailItem" Then
        Debug.print(Item.SenderEmailAddress)
        If Item.Subject == "PADRAO" Then
            'CODE HERE
        End If
    End If


Except_capture:
    Debug.print("Houve um erro no evento de captura. "+ vbNewLine +_
        "Código do erro: " + str(Err.Number) + " e descricão: " + Err.Description)
    Exit Sub
End Sub

Sub empilha_excel(ByVal Item as Outlook.MailItem)
    On Error GoTo Except_append
    Set myattachments = Item.Attachments 
    
    While myattachments.Count > 0 
        myattachments.Item(1).FileName
        'myattachments.Item(1).SaveAsFile("")
    Wend 

Except_append:
    Debug.print("Houve um erro no metodo de empilhar. "+ vbNewLine +_
        "Código do erro: " + str(Err.Number) + " e descricão: " + Err.Description)
End Sub

Sub enviaEmail(tolist As String, bodyText As String, subjectext As String, Optional copyList As String, Optional pathfile As String = "")
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim texto As String
    Dim mensagem As String
    'Define o objeto Outlook numa variável
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp. CreateItem(0)
    On Error Resume Next
    
    With OutMail
        .To = toList
        .CC = copyList
        .Subject = subjectText
        _HTMLBody = bodyText
        If anexo <> "" Then
            .Attachments.Add anexo
        End If
        '.Display
        . Send
    End With
    
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub