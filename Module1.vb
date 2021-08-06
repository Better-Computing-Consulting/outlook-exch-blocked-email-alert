Imports Microsoft.Office.Interop
Imports System.Net.Mail
Module Module1
    Sub Main()
        Dim OlApp As New Outlook.Application
        Dim oNS As Outlook.NameSpace = OlApp.GetNamespace("MAPI")
        Dim ibox As Outlook.MAPIFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        AddHandler OlApp.NewMailEx, AddressOf NewMailEx_Handler
        Console.WriteLine("Waiting for emails...")
        Console.ReadLine()
    End Sub
    Sub NewMailEx_Handler(ByVal EntryIDCollection As String)
        Dim tmpIDs As String = EntryIDCollection
        For Each sID As String In tmpIDs.Split(",")
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & ": " & sID)
            ProcessOneEmail(sID)
        Next
    End Sub
    Sub ProcessOneEmail(emailid As String)
        Dim OlApp As New Outlook.Application
        Dim oNS As Outlook.NameSpace = OlApp.GetNamespace("MAPI")
        Dim ibox As Outlook.MAPIFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        oNS.Logon()
        Try
            Dim i As Outlook.MailItem = oNS.GetItemFromID(emailid)
            If i.Attachments.Count > 0 Then
                Dim sAttachments As New List(Of String)
                For Each atch As Outlook.Attachment In i.Attachments
                    If atch.FileName.ToLower.EndsWith(".zip") Then
                        sAttachments.Add(atch.FileName)
                    End If
                Next
                If sAttachments.Count > 0 Then
                    Dim sSender As String = ""
                    Dim sReceivers As New List(Of String)
                    Dim sSubject As String = ""
                    Dim sBody As String = ""
                    If i.SenderEmailType = "EX" Then
                        Dim sender As Outlook.AddressEntry = i.Sender
                        sSender = GetExchangeSMTPAddress(sender)
                    Else
                        sSender = i.SenderName & "<" & i.SenderEmailAddress & ">"
                    End If
                    For Each r As Outlook.Recipient In i.Recipients
                        If r.AddressEntry.Type = "EX" Then
                            Dim tempReceiver As String = GetExchangeSMTPAddress(r.AddressEntry)
                            If tempReceiver.ToLower.Contains("@lawfirm.com") Then
                                sReceivers.Add(tempReceiver)
                            End If
                        End If
                    Next
                    sSubject = i.Subject
                    sBody = i.Body
                    Console.WriteLine("  From: " & sSender)
                    Console.WriteLine("  To  : " & String.Join(",", sReceivers.ToArray))
                    Console.WriteLine("  Subj: " & sSubject)
                    Console.WriteLine("  Atch: " & String.Join(",", sAttachments.ToArray))
                    SendEmail(sSender, sReceivers, sSubject, sBody, sAttachments, OlApp.CreateItem(Outlook.OlItemType.olMailItem))
                    Dim dirID As String = ""
                    Dim folderfound As Boolean = False
                    For Each fld As Outlook.MAPIFolder In ibox.Folders
                        If fld.Name = "Processed" Then
                            folderfound = True
                            Try
                                i.Move(fld)
                                Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & ": Message moved to Processed directory.")
                            Catch ex2 As Exception
                                Console.WriteLine(ex2.Message)
                            End Try
                        End If
                    Next
                    If Not folderfound Then
                        Try
                            Dim fld As Outlook.MAPIFolder = ibox.Folders.Add("Processed", Outlook.OlDefaultFolders.olFolderInbox)
                            i.Move(fld)
                            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & ": Message moved to New Processed directory.")
                        Catch ex2 As Exception
                            Console.WriteLine(ex2.Message)
                        End Try
                    End If
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
    Sub SendEmail(inFrom As String, inTo As List(Of String), inSubject As String, inBody As String, inAttachments As List(Of String), olMsg As Outlook.MailItem)
        Dim sMsg As Outlook.MailItem
        sMsg = olMsg
        Dim tmpSubj As String = "This message from " & inFrom & " was blocked because it contained zip file(s) " & String.Join(", ", inAttachments.ToArray) & ". " & inFrom & " received a bounce message stating that the law firm does not accept ZIP files by E-mail."
        tmpSubj &= vbCrLf & vbCrLf & "If you are sure the email is valid, the zip file is safe and expected, and cannot wait for the other party to resend it, please contact IT."
        tmpSubj &= vbCrLf & vbCrLf & "====================================================================================" & vbCrLf & vbCrLf
        With sMsg
            For Each s As String In inTo
                .Recipients.Add(s)
            Next
            Dim olRec As Outlook.Recipient
            olRec = .Recipients.Add("it@lawfirm.com")
            olRec.Type = Outlook.OlMailRecipientType.olCC
            For Each r As Outlook.Recipient In .Recipients
                r.Resolve()
            Next
            .Subject = "[ZIP BLOCKED] " & inSubject
            .Body = tmpSubj & inBody
            .BodyFormat = Outlook.OlBodyFormat.olFormatHTML
        End With
        Try
            sMsg.Send()
            Console.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & ": Informational email to blocked recipients sent.")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
    Sub CheckEmails()
        Dim OlApp As New Outlook.Application
        Dim oNS As Outlook.NameSpace = OlApp.GetNamespace("MAPI")
        Dim ibox As Outlook.MAPIFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        oNS.Logon()
        For Each i As Outlook.MailItem In ibox.Items
            If i.Attachments.Count > 0 Then
                Dim sSender As String = ""
                Dim sReceivers As New List(Of String)
                Dim sSubject As String = ""
                Dim sBody As String = ""
                Dim sAttachments As New List(Of String)

                If i.SenderEmailType = "EX" Then
                    Dim sender As Outlook.AddressEntry = i.Sender
                    sSender = GetExchangeSMTPAddress(sender)
                Else
                    sSender = i.SenderEmailAddress
                End If
                For Each r As Outlook.Recipient In i.Recipients
                    Console.WriteLine(">>>> " & r.AddressEntry.Type & " " & r.Address)
                    If r.AddressEntry.Type = "EX" Then
                        Console.WriteLine(">>>> " & r.AddressEntry.AddressEntryUserType)
                        Dim tempReceiver As String = GetExchangeSMTPAddress(r.AddressEntry)
                        Console.WriteLine(">>>>>" & tempReceiver)
                        If tempReceiver.ToLower.Contains("@lawfirm.com") Then
                            sReceivers.Add(tempReceiver)
                        End If
                    End If
                Next
                sSubject = i.Subject
                For Each atch As Outlook.Attachment In i.Attachments
                    If atch.FileName.ToLower.EndsWith(".txt") Then
                        sAttachments.Add(atch.FileName)
                    End If
                Next
                sBody = i.Body
                Console.WriteLine(sSender)
                For Each r As String In sReceivers
                    Console.WriteLine(r)
                Next
                Console.WriteLine(sSubject)
                For Each a As String In sAttachments
                    Console.WriteLine(a)
                Next
                Console.WriteLine("####################################################")
                Console.WriteLine(sBody)
                Console.WriteLine("####################################################")
            End If
        Next
    End Sub
    Function GetExchangeSMTPAddress(inaddr As Outlook.AddressEntry) As String
        Dim result As String = ""
        Dim addr As Outlook.AddressEntry = inaddr
        If addr IsNot Nothing Then
            If addr.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry OrElse addr.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
                Dim exaddr As Outlook.ExchangeUser = addr.GetExchangeUser
                If exaddr IsNot Nothing Then
                    result = exaddr.PrimarySmtpAddress
                End If
            ElseIf addr.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry Then
                Dim exdist As Outlook.ExchangeDistributionList = addr.GetExchangeDistributionList
                If exdist IsNot Nothing Then
                    result = exdist.PrimarySmtpAddress
                End If
            End If
        End If
        Return result
    End Function
End Module
