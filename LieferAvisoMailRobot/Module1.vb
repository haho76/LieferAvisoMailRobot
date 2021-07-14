Imports System.Data.SqlClient
'Imports Outlook = Microsoft.Office.Interop.Outlook


Module Module1

    Sub Main()
        'Dim oOutlook As New Outlook.Application
        Dim oOutlook As Object
        'Dim oNs As Outlook.NameSpace
        Dim oNs As Object
        'Dim oFldr As Outlook.MAPIFolder
        Dim oFldr As Object
        Dim oMessage As Object
        'Dim oForwardMessage As Outlook.MailItem
        'Dim oCopy As Outlook.MailItem
        Dim oForwardMessage As Object
        Dim oCopy As Object
        Dim sMailRecepient As String
        Dim sSubjects() As String = Nothing 'Speichert gefundene Treffer an Betreffen
        Dim sSMTPAdress As String = Nothing
        Dim iSubjectCounter As Integer = 0 'Array Counter
        Dim bMatchExists As Boolean = False

        Try
            'Latebinding
            oOutlook = CreateObject("Outlook.Application")
            'Folder setzen
            oNs = oOutlook.GetNamespace("MAPI")
            oFldr = oNs.GetDefaultFolder(oOutlook.OlDefaultFolders.olFolderInbox)

            'Mails durchgehen - wenn Stichwort gefunden Adresse auslesen und zugehörige Mailadresse abfragen
            For Each oMessage In oFldr.Items
                If oMessage.Class = oOutlook.OlObjectClass.olMail Then   'Wenn Objekt = Mailitem (Report Items würden sonst Fehler verursachen)
                    'Case: Betreff enthält "Lieferbenachrichtigung" oder "Abliefernachweis" und kommt von "AutoMail@ingrammicro.de"
                    If (oMessage.Subject Like "*Lieferbenachrichtigung*" Or oMessage.subject Like "*Abliefernachweis*") And Get_SMTP(oMessage, oOutlook) = "IngramMicroAutomail@ingrammicro.com" Then
                        bMatchExists = True                                    'Boolean auf True setzen - Treffer gefunden
                        ReDim Preserve sSubjects(iSubjectCounter)              'Array erweitern
                        sSubjects(iSubjectCounter) = oMessage.Subject.ToString 'Betreff in Array schreiben

                        'Mail erstellen/senden
                        oForwardMessage = oMessage.Forward   'Erstellt Forwardmail
                        sMailRecepient = Get_Mailadress(Get_AdressData(oMessage.Body.Split(vbCrLf))) 'Mailadresse anhand Lieferadresse auslesen

                        With oForwardMessage
                            'Wenn Adresse gefunden, Mail verschicken, wenn nicht Mail an AT IT und Mail verschieben
                            If Not sMailRecepient = Nothing Then
                                .To = sMailRecepient
                                .Subject = oMessage.Subject  'Betreff = Originalbetreff
                                .BCC = "harald.holzer@ingrammicro.com"
                            Else
                                .To = "IT.I@ingrammicro.at"
                                .Subject = "Kein Eintrag in Tabelle! /" & oMessage.Subject
                                oCopy = .Copy  'Mailkopie erstellen / Original wird später gelöscht
                                oCopy.Move(oFldr.Folders(My.Settings.sOLSubfolderError))  'In Fehlerordner verschieben
                            End If
                            .HTMLBody = oMessage.HTMLBody  'Body = Originalbody
                            .Send()
                        End With
                        iSubjectCounter = iSubjectCounter + 1  'Zähler erhöhen
                        'Case: Betreff enthält "Lieferbenachrichtigung" oder "Abliefernachweis" und kommt NICHT von "AutoMail@ingrammicro.de" -> Antworten behandeln
                    ElseIf (oMessage.Subject Like "*Lieferbenachrichtigung*" Or oMessage.To = "Lieferaviso@ingrammicro.at") And Get_SMTP(oMessage, oOutlook) <> "IngramMicroAutomail@ingrammicro.com" Then
                        bMatchExists = True                         'Boolean auf True setzen - Treffer gefunden
                        ReDim Preserve sSubjects(iSubjectCounter)    ' Array erweitern
                        sSubjects(iSubjectCounter) = oMessage.Subject.ToString 'Betreff in array schreiben

                        'Mail erstellen/senden
                        oForwardMessage = oMessage.Forward
                        With oForwardMessage
                            .To = "ecomm@ingrammicro.at"
                            .Subject = oMessage.Subject  'Betreff = Originalbetreff
                            .Send()
                        End With
                        iSubjectCounter = iSubjectCounter + 1  'Zähler erhöhen
                    End If
                End If
            Next

            ''Gefundene Mails löschen
            'If bMatchExists Then
            '    Dim oMailToDel As Outlook.MailItem
            '    Dim iMailItemCounter As Integer
            '    'Durch gespeicherte BEtreffs gehen und entsprechende Mails löschen
            '    For iSubjectCounter = 0 To UBound(sSubjects)
            '        iMailItemCounter = 1
            '        While iMailItemCounter <= oFldr.Items.Count
            '            If oFldr.Items(iMailItemCounter).class = Outlook.OlObjectClass.olMail Then
            '                oMailToDel = oFldr.Items(iMailItemCounter)  'Mail das gelöscht werden soll setzen
            '                If oMailToDel.Subject = sSubjects(iSubjectCounter) Then  'prüfen ob Betreff richtig
            '                    oMailToDel.Delete()
            '                    Exit While
            '                End If
            '            End If
            '            iMailItemCounter = iMailItemCounter + 1
            '        End While
            '    Next
            '    oMailToDel = Nothing
            'End If

            ''Cleanup
            'oOutlook = Nothing
            'oNs = Nothing
            'oFldr = Nothing
            'oMessage = Nothing
            'oForwardMessage = Nothing
            'oCopy = Nothing

            'SQL for Tracking Tool
            Dim sqlConnection1 As New SqlConnection(My.Settings.ConnectionString)
            Dim cmd As New SqlCommand
            Dim returnValue As Object
            cmd.CommandText = "Update [APPSEC].[dbo].[TrackingTool] set Last_Update = getdate() where Application_ID = 6"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1
            sqlConnection1.Open()
            returnValue = cmd.ExecuteNonQuery
            sqlConnection1.Close()

        Catch ex As Exception
            'Dim olErrMail As Outlook.MailItem
            Dim olErrMail As Object
            olErrMail = oOutlook.CreateItem(oOutlook.OlItemType.olMailItem)
            With olErrMail
                .To = "IT.I@ingrammicro.at"
                .Subject = "Lieferaviso-Fehler"
                .Body = ex.Message.ToString
                .Send()
            End With
            olErrMail = Nothing
        Finally
            'Gefundene Mails löschen
            If bMatchExists Then
                'Dim oMailToDel As Outlook.MailItem
                Dim oMailToDel As Object
                Dim iMailItemCounter As Integer
                'Durch gespeicherte BEtreffs gehen und entsprechende Mails löschen
                For iSubjectCounter = 0 To UBound(sSubjects)
                    iMailItemCounter = 1
                    While iMailItemCounter <= oFldr.Items.Count
                        If oFldr.Items(iMailItemCounter).class = oOutlook.OlObjectClass.olMail Then
                            oMailToDel = oFldr.Items(iMailItemCounter)  'Mail das gelöscht werden soll setzen
                            If oMailToDel.Subject = sSubjects(iSubjectCounter) Then  'prüfen ob Betreff richtig
                                oMailToDel.Delete()
                                Exit While
                            End If
                        End If
                        iMailItemCounter = iMailItemCounter + 1
                    End While
                Next
                oMailToDel = Nothing
            End If

            'Cleanup
            oOutlook = Nothing
            oNs = Nothing
            oFldr = Nothing
            oMessage = Nothing
            oForwardMessage = Nothing
            oCopy = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Function reads through mailbody and returns Customerdata found based on keywords
    ''' </summary>
    ''' <param name="strMailBody">Mailbody String Array</param>
    ''' <returns>Concated string with Customer/Adressdata</returns>
    ''' <remarks></remarks>
    Public Function Get_AdressData(strMailBody() As String) As String
        Dim sCustNo As String = Nothing
        Dim sAdress1 As String = Nothing
        Dim sAdress2 As String = Nothing
        Dim sAdress3 As String = Nothing
        Dim sZipCity As String = Nothing
        Dim Counter As Integer

        'Go through String Array
        For i = 0 To UBound(strMailBody)
            'Find Customernumber based on keyword
            If strMailBody(i).IndexOf(My.Settings.CustNoSearchKey) > -1 Then
                If Len(strMailBody(i)) > Len("Kundennummer:") Then
                    sCustNo = strMailBody(i).Replace("Kundennummer:", "").Trim.Replace("-", "")
                Else
                    For j = i To UBound(strMailBody) 'For j = i + 1 To UBound(strMailBody)
                        If Len(strMailBody(j)) >= 9 Then
                            sCustNo = strMailBody(j).Trim().Replace("-", "")
                            Exit For
                        End If
                    Next
                End If
            End If
                'Find Adressdata based on keyword
                If strMailBody(i).IndexOf(My.Settings.AdressSearchKeyWord) > -1 Then
                    sAdress1 = System.Text.RegularExpressions.Regex.Replace(strMailBody(i + 1), "<.*?>", "").Trim 'Löscht HTML Tags in Kundennamen (zb. Mobile.com)
                    sAdress1 = System.Text.RegularExpressions.Regex.Replace(sAdress1, "  +", " ") 'Kürzt mehrfache Leerzeichen zu einem  
                    sAdress2 = strMailBody(i + 2).Trim
                    sAdress3 = strMailBody(i + 3).Trim
                    sZipCity = strMailBody(i + 4).Trim
                    Counter = i
                    Exit For
                End If
        Next
        'Return concated string
        Return sCustNo & sAdress1 & sAdress2 & sAdress3 & sZipCity
    End Function


    ''' <summary>
    ''' Looksup mail adress according to deliveryadress/suffix
    ''' </summary>
    ''' <param name="sAdressData">Concated AdressData</param>
    ''' <returns>mailadress</returns>
    ''' <remarks></remarks>
    Public Function Get_Mailadress(sAdressData As String) As String
        Dim sqlConnection1 As New SqlConnection(My.Settings.ConnectionString)
        Dim cmd As New SqlCommand
        Dim returnValue As Object


        cmd.CommandText = "SELECT eMail FROM [Tools].[dbo].[Lieferbestätigung_an_Suffix] WHERE RTRIM(CUST_NO)+RTRIM(Suffix) = " &
                               "(SELECT RTRIM(branch_customer_nbr)+RTRIM(suffix) FROM [DSS001].[dbo].[dss_customer_location] " &
                               "WHERE RTRIM([branch_customer_nbr])+RTRIM([cust_add1])+RTRIM([cust_add2])+RTRIM([cust_add3])+RTRIM([cust_city]) = '" & sAdressData & "')"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1
        sqlConnection1.Open()
        returnValue = cmd.ExecuteScalar()

        'wenn kein Ergebnis - nimm 000 Suffix
        If IsNothing(returnValue) Then
            cmd.CommandText = "SELECT eMail FROM [Tools].[dbo].[Lieferbestätigung_an_Suffix] WHERE RTRIM(CUST_NO)+RTRIM(Suffix) = '" & Left(sAdressData, 8) & "000" & "'"
            returnValue = cmd.ExecuteScalar()
        End If

        sqlConnection1.Close()

        Return returnValue
    End Function

    ''' <summary>
    ''' Liest Mailadresse aus Email aus - egal ob Exchange oder SMTP
    ''' </summary>
    ''' <param name="mail">Mailitem</param>
    ''' <returns>MAiladresse als string</returns>
    ''' <remarks></remarks>
    Public Function Get_SMTP(mail As Object, oOutlook As Object) As String
        Dim sender As Object
        Dim exchUser As Object

        sender = mail.Sender
        'Prüft ob Exchange Mail oder SMTP
        If sender.AddressEntryUserType = oOutlook.OlAddressEntryUserType.olExchangeUserAddressEntry Or sender.AddressEntryUserType = oOutlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
            'Wenn Exchange, GEt User & Return Primary SMTP Adress
            exchUser = sender.GetExchangeUser
            If exchUser IsNot Nothing Then
                Return exchUser.PrimarySmtpAddress
            Else : Return String.Empty
            End If
            'SMPT -> Return SMTP (SenderEmailAdress)
        Else
            Return mail.SenderEmailAddress
        End If
    End Function

End Module
