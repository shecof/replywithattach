
Sub forward_mail(attachments, item) 'Ответ на письмо с вложением
    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    Dim olRecip As Recipient
    
        Set olReply = item.ReplyAll
        item.UnRead = False
        
    For i = 1 To UBound(attachments)
        If Not attachments(i) = "" Then
            olReply.attachments.Add (attachments(i))
        End If
    Next i
        olReply.Display
        

End Sub


Sub Prepare_attachments(windowtype As Integer) 'сохранение файлов из письма в папку
    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    Dim olRecip As Recipient
    Dim myinspector As Outlook.Inspector
    Dim mail_attachments As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    
    If windowtype = 1 Then
        If Application.ActiveInspector Is Nothing Then
            Set item = Application.ActiveExplorer.Selection.item(1)
        Else
            Set item = Application.ActiveExplorer.Selection.item(1)
        End If
    Else
        If Application.ActiveInspector Is Nothing Then
            Set item = Application.ActiveExplorer.Selection.item(1)
        Else
            Set item = Application.ActiveInspector.CurrentItem
        End If
    End If
    

    mail_path = VBA.Environ$("USERPROFILE") & "\OutlookAttachments\"
    
    On Error Resume Next
    fso.DeleteFile mail_path & "*"

    Dim attach_title As String
    Dim temp_attach_title() As String
    Dim attachments_array() As String
    ReDim attachments_array(100) 'максимальное количество файлов вложений
    
    
    If item.attachments.Count > 0 Then
        For j = 1 To item.attachments.Count
            If Not item.attachments.item(j).FileName Like "*image0*" Then 'пропуск картинок из подписей
                attach_title = item.attachments.item(j).FileName
                If item.attachments.item(j).FileName Like "RE *" Then
                    temp_attach_title() = Split(item.attachments.item(j).FileName, "RE ") 'переименование пересланных и отвеченных писем
                    attach_title = temp_attach_title(1)
                ElseIf item.attachments.item(j).FileName Like "FW *" Then
                    temp_attach_title() = Split(item.attachments.item(j).FileName, "FW ")
                    attach_title = temp_attach_title(1)
                ElseIf item.attachments.item(j).FileName Like "FW: *" Then
                    temp_attach_title() = Split(item.attachments.item(j).FileName, "FW: ")
                    attach_title = temp_attach_title(1)
                ElseIf item.attachments.item(j).FileName Like "RE: *" Then
                    temp_attach_title() = Split(item.attachments.item(j).FileName, "RE: ")
                    attach_title = temp_attach_title(1)
                End If
                
                path = mail_path & j & "_" & attach_title
                item.attachments.item(j).SaveAsFile path
                attachments_array(j) = path
            End If
        Next j
    End If
    

    
    forward_mail attachments_array(), item
    
End Sub

Sub btn_reply_explorer() 'Вызов функции из explorer (используется в кнопке)
    Prepare_attachments (1)
End Sub

Sub btn_reply_inspector() 'Вызов функции из inspector (используется в кнопке)
    Prepare_attachments (2)
End Sub
