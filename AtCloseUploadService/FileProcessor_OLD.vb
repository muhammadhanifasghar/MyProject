﻿Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Xml
Imports System.Net.Mail
Imports System.Configuration
Imports ReswareUploader
Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class FileProcessor
    Private strClientNumber As String
    Private strLoanNumber As String
    Private strOrderNumber As String

    Enum ErrorType
        Retry
        InvalidId
        IncorrectId
        System
    End Enum

    Public Function ProcessFile(strFile As String) As String
        'Returns OK if no errors. Error message if appropriate
        Dim product As S180Product = New S180Product
        Dim document As S180Document = New S180Document
        Dim documentList As List(Of S180Document) = New List(Of S180Document)
        Dim productList As List(Of S180Product) = New List(Of S180Product)
        Dim S180event As S180Event = New S180Event
        Dim strSplit() As String
        Dim strStatus As String = "Error"
        Dim strStatusVisi As String = "Error"
        Dim strStatusWFG As String = "Error"
        Dim strStatusResware As String = "Error"
        Dim strStatusSilkResware As String = "Error"
        Dim strStatusCommonwealth As String = "Error"
        Dim strStatusMessage As StringBuilder = New StringBuilder
        Dim oErrorType As ErrorType
        Dim strDBStatus As String
        Dim blnIsVSIRetry As Boolean = False
        Dim blnIsWFGRetry As Boolean = False
        Dim blnIsReswareRetry As Boolean = False
        Dim blnIsCommonwealthRetry As Boolean = False
        Dim blnIsSilkReswareRetry As Boolean = False

        Try
            strSplit = strFile.Split(New Char() {"_"c})
            oServiceLog.WriteLogEntry("Client Code = " + strSplit(1))
            oServiceLog.WriteLogEntry("Loan Number = " + strSplit(3))
            oServiceLog.WriteLogEntry("Order Number = " + strSplit(2))

            product.UniqueId = "1"
            product.Instructions = ""
            product.Reference = ""
            product.ProviderOrderNbr = Path.GetFileNameWithoutExtension(strSplit(2))
            S180event.Code = "S180"
            S180event.EventDate = String.Format(Date.Now, "yyyy/MM/dd HH:mm:ss")
            S180event.OccurDate = ""
            S180event.ClientCode = ""
            S180event.ClientDesc = ""
            S180event.Comment = "Please find attached a supporting document for this order."
            S180event.EventId = DateTime.Now.ToString("yyMMddHHmmss")
            product.S180Event = S180event
            document.DocDate = File.GetLastWriteTime(strFile)
            document.DocVersion = "1"
            document.DocStatus = "FINAL"
            document.DocName = Path.GetFileName(strFile)
            document.FileName = Path.GetFileName(strFile)
            document.FileType = "PDF"
            document.ExpireDate = ""
            document.Content = ConvertFileToBase64(strFile)
            document.EncodeType = "Base64"

#If False Then

            ' --
            ' -- Neg 3-28-2016
            ' --
            ' -- This is the code that was changed on 3-28 for the new jobs on the IBML scanners. The IF statement was replaced with the commented one below

            If strFile.Contains("Recordable") Then
                document.DocType = "RECORDABLE"
                document.Description = "Recording Documents"
            Else
                document.DocType = "POSTCLDOC"
                document.Description = "Signed-scanned Closing Package"
            End If
#End If
            If strFile.Contains("Recordable") Then
                product.UniqueId = "1"   ' Closing Product
                document.DocType = "RECORDABLE"
                document.Description = "Recording Documents"
            ElseIf strFile.Contains("Unrecorded") Then
                product.UniqueId = "3"     ' Recording Product
                document.DocType = "RECORDABLE"
                document.Description = "Recording Documents"
            ElseIf strFile.Contains("Recorded") Then
                product.UniqueId = "3" ' Recording Product
                document.DocType = "RECORDED"
                document.Description = "Recorded Documents"
            Else
                product.UniqueId = "1"  ' Closing
                document.DocType = "POSTCLDOC"
                document.Description = "Signed-scanned Closing Package"
            End If

            document.DocTypeOtherDesc = ""
            product.DocumentList = documentList
            product.DocumentList.Add(document)

            Dim outputFile As S180XMLFile = New S180XMLFile
            outputFile.TransactionId = Path.GetFileNameWithoutExtension(strSplit(2))
            outputFile.ProductList = productList
            outputFile.ProductList.Add(product)

            If Path.GetFileNameWithoutExtension(strFile).Substring(0, 3) = "111" Then
                blnUploadToVSI = True
                blnIsVSIRetry = True
                blnUploadToWFG = True
                blnUploadToCommonwealth = True
                blnIsWFGRetry = True
                blnUploadToResware = True
                blnIsReswareRetry = True
                blnIsCommonwealthRetry = True
            ElseIf Path.GetFileNameWithoutExtension(strFile).Substring(0, 1) = "1" Then
                blnUploadToVSI = True
                blnIsVSIRetry = True
                blnUploadToWFG = False
                blnUploadToResware = False
                blnUploadToCommonwealth = False
                blnUploadToSilkResware = False
                strStatusResware = "OK"
                strStatusWFG = "OK"
                strStatusCommonwealth = "OK"
                strStatusSilkResware = "OK"
            ElseIf Path.GetFileNameWithoutExtension(strFile).Substring(1, 1) = "1" Then
                blnUploadToWFG = True
                blnIsWFGRetry = True
                blnUploadToVSI = False
                blnUploadToResware = False
                blnUploadToCommonwealth = False
                blnUploadToSilkResware = False
                strStatusVisi = "OK"
                strStatusResware = "OK"
                strStatusCommonwealth = "OK"
                strStatusSilkResware = "OK"

            ElseIf Path.GetFileNameWithoutExtension(strFile).Substring(2, 1) = "1" Then
                blnUploadToResware = True
                blnIsReswareRetry = True
                blnUploadToVSI = False
                blnUploadToWFG = False
                blnUploadToCommonwealth = False
                blnUploadToSilkResware = False
                strStatusVisi = "OK"
                strStatusWFG = "OK"
                strStatusCommonwealth = "OK"
                strStatusSilkResware = "OK"
            ElseIf Path.GetFileNameWithoutExtension(strFile).Substring(3, 1) = "1" Then
                blnUploadToCommonwealth = True
                blnIsCommonwealthRetry = True
                blnUploadToVSI = False
                blnUploadToWFG = False
                blnUploadToResware = False
                blnUploadToSilkResware = False
                strStatusVisi = "OK"
                strStatusWFG = "OK"
                strStatusResware = "OK"
                strStatusSilkResware = "OK"
            ElseIf Path.GetFileNameWithoutExtension(strFile).Substring(4, 1) = "1" Then
                blnUploadToCommonwealth = False
                blnUploadToVSI = False
                blnUploadToWFG = False
                blnUploadToResware = False
                blnUploadToSilkResware = True
                blnIsSilkReswareRetry = True
                strStatusVisi = "OK"
                strStatusWFG = "OK"
                strStatusResware = "OK"
                strStatusSilkResware = "OK"
            Else
                'Get list of clients to determine if this file should be uploaded to WFG as well as Visionet
                blnUploadToVSI = True
                Dim xmlFile As XmlReader
                oServiceLog.WriteLogEntry("ClientXMLFile = " + System.Configuration.ConfigurationManager.AppSettings("ClientXMLFile"))
                xmlFile = XmlReader.Create(System.Configuration.ConfigurationManager.AppSettings("ClientXMLFile"), New XmlReaderSettings)
                Dim ds As New DataSet
                Dim dv As DataView
                ds.ReadXml(xmlFile)
                dv = New DataView(ds.Tables(0))
                Try
                    dv.Sort = "ClientCode_Text"
                Catch
                    dv.Sort = "ClientCode"
                End Try
                Dim index As Integer = dv.Find(strSplit(1))
                If index <> -1 Then
                    blnUploadToWFG = True
                Else
                    blnUploadToWFG = False
                    strStatusWFG = "OK"
                End If
                Dim xmlFileResware As XmlReader
                oServiceLog.WriteLogEntry("ReswareClientXMLFile = " + System.Configuration.ConfigurationManager.AppSettings("ReswareClientXMLFile"))
                xmlFileResware = XmlReader.Create(System.Configuration.ConfigurationManager.AppSettings("ReswareClientXMLFile"), New XmlReaderSettings)
                Dim dsResware As New DataSet
                Dim dvResware As DataView
                dsResware.ReadXml(xmlFileResware)
                dvResware = New DataView(dsResware.Tables(0))
                Try
                    dvResware.Sort = "ClientCode_Text"
                Catch
                    dvResware.Sort = "ClientCode"
                End Try
                Dim indexResware As Integer = dvResware.Find(strSplit(1))
                If indexResware <> -1 And Not strFile.Contains("Recordable") Then
                    blnUploadToResware = True
                Else
                    blnUploadToResware = False
                    strStatusResware = "OK"
                End If
                Dim xmlFileSilkResware As XmlReader
                oServiceLog.WriteLogEntry("SilkReswareClientXMLFile = " + System.Configuration.ConfigurationManager.AppSettings("SilkReswareClientXMLFile"))
                xmlFileSilkResware = XmlReader.Create(System.Configuration.ConfigurationManager.AppSettings("SilkReswareClientXMLFile"), New XmlReaderSettings)
                Dim dsSilkResware As New DataSet
                Dim dvSilkResware As DataView
                dsSilkResware.ReadXml(xmlFileSilkResware)
                dvSilkResware = New DataView(dsSilkResware.Tables(0))
                Try
                    dvSilkResware.Sort = "ClientCode_Text"
                Catch
                    dvSilkResware.Sort = "ClientCode"
                End Try
                Dim indexSilkResware As Integer = dvSilkResware.Find(strSplit(1))
                If indexSilkResware <> -1 And Not strFile.Contains("Recordable") Then
                    blnUploadToSilkResware = True
                Else
                    blnUploadToSilkResware = False
                    strStatusSilkResware = "OK"
                End If
                Dim xmlFileCommonwealth As XmlReader
                oServiceLog.WriteLogEntry("CommonwealthXMLFile = " + System.Configuration.ConfigurationManager.AppSettings("CommonwealthXMLFile"))
                xmlFileCommonwealth = XmlReader.Create(System.Configuration.ConfigurationManager.AppSettings("CommonwealthXMLFile"), New XmlReaderSettings)
                Dim dsCommonwealth As New DataSet
                Dim dvCommonwealth As DataView
                dsCommonwealth.ReadXml(xmlFileCommonwealth)
                dvCommonwealth = New DataView(dsCommonwealth.Tables(0))
                Try
                    dvCommonwealth.Sort = "ClientCode_Text"
                Catch
                    dvCommonwealth.Sort = "ClientCode"
                End Try
                Dim indexCommonwealth As Integer = dvCommonwealth.Find(strSplit(1))
                If indexCommonwealth <> -1 Then
                    blnUploadToCommonwealth = True
                Else
                    blnUploadToCommonwealth = False
                    strStatusCommonwealth = "OK"
                End If
            End If

            oServiceLog.WriteLogEntry("Upload to VSI = " + blnUploadToVSI.ToString)
            oServiceLog.WriteLogEntry("Upload to WFG = " + blnUploadToWFG.ToString)
            oServiceLog.WriteLogEntry("Upload to Resware = " + blnUploadToResware.ToString)
            oServiceLog.WriteLogEntry("Upload to Commonwealth = " + blnUploadToCommonwealth.ToString)
            oServiceLog.WriteLogEntry("Upload to SilkReswer = " + blnUploadToSilkResware.ToString)

            oServiceLog.WriteLogEntry("Creating Visionet XML file to: " + strTempDir + "\" + Path.GetFileNameWithoutExtension(strFile) + ".xml")
            strStatus = outputFile.WriteToFile(strTempDir + "\" + Path.GetFileNameWithoutExtension(strFile) + ".xml")
            oServiceLog.WriteLogEntry("Visionet XML file created")

            product.ProviderOrderNbr = Path.GetFileNameWithoutExtension(strSplit(3))
            outputFile.TransactionId = Path.GetFileNameWithoutExtension(strSplit(3))
            oServiceLog.WriteLogEntry("Creating WFG XML file to: " + strTempDir + "\WFG" + Path.GetFileNameWithoutExtension(strFile) + ".xml")
            strStatus = outputFile.WriteToFile(strTempDir + "\WFG" + Path.GetFileNameWithoutExtension(strFile) + ".xml")
            oServiceLog.WriteLogEntry("WFG XML file created")

            oServiceLog.WriteLogEntry("Creating Commonwealth XML file to: " + strTempDir + "\Commonwealth" + Path.GetFileNameWithoutExtension(strFile) + ".xml")
            strStatus = outputFile.WriteToFile(strTempDir + "\Commonwealth" + Path.GetFileNameWithoutExtension(strFile) + ".xml")
            oServiceLog.WriteLogEntry("Commonwealth XML file created")

            If blnUploadToWFG Then
                objScanningRecord = New ScanningRecord
                oServiceLog.WriteLogEntry("Uploading file " + strFile + " to WFG")
                objScanningRecord.UploadStart = Now.ToString
                strStatusWFG = UploadXMLFile(strTempDir + "\WFG" + Path.GetFileNameWithoutExtension(strFile) + ".xml", strWFGURL)
                objScanningRecord.UploadEnd = Now.ToString
                If strStatusWFG = "OK" Then
                    objScanningRecord.Uploaded = 1
                Else
                    objScanningRecord.Exception = 1
                    objScanningRecord.Message = strStatusWFG
                    oErrorType = GetErrorType(strStatusWFG)
                    Select Case oErrorType
                        Case ErrorType.Retry
                            objScanningRecord.TimeOut = 1
                        Case ErrorType.IncorrectId
                            objScanningRecord.IncorrectId = 1
                        Case ErrorType.InvalidId
                            objScanningRecord.InvalidId = 1
                    End Select
                End If
                objScanningRecord.SystemType = "WFG"
                objScanningRecord.ClientCode = strSplit(1)
                objScanningRecord.LoanNumber = strSplit(3)
                objScanningRecord.OrderNumber = strSplit(2)
                objScanningRecord.FileName = Path.GetFileName(strFile)
                objScanningRecord.MachineName = Environment.MachineName
                If blnIsWFGRetry Then
                    objScanningRecord.RetryAttempts = GetRetryAttempts(strFile) + 1
                End If
                strDBStatus = objScanningRecord.AddScanningRecord()
                If strDBStatus <> "OK" Then
                    SendNotificationEmail("Error adding record to database for file WFG: " + strFile, strDBStatus, "ERR14")
                End If
            End If

            If blnUploadToCommonwealth Then
                objScanningRecord = New ScanningRecord
                oServiceLog.WriteLogEntry("Uploading file " + strFile + " to Commonwealth")
                objScanningRecord.UploadStart = Now.ToString
                strStatusCommonwealth = UploadXMLFile(strTempDir + "\Commonwealth" + Path.GetFileNameWithoutExtension(strFile) + ".xml", strCommonwealthURL)
                objScanningRecord.UploadEnd = Now.ToString
                If strStatusCommonwealth = "OK" Then
                    objScanningRecord.Uploaded = 1
                Else
                    objScanningRecord.Exception = 1
                    objScanningRecord.Message = strStatusCommonwealth
                    oErrorType = GetErrorType(strStatusCommonwealth)
                    Select Case oErrorType
                        Case ErrorType.Retry
                            objScanningRecord.TimeOut = 1
                        Case ErrorType.IncorrectId
                            objScanningRecord.IncorrectId = 1
                        Case ErrorType.InvalidId
                            objScanningRecord.InvalidId = 1
                    End Select
                End If
                objScanningRecord.SystemType = "Commonwealth"
                objScanningRecord.ClientCode = strSplit(1)
                objScanningRecord.LoanNumber = strSplit(3)
                objScanningRecord.OrderNumber = strSplit(2)
                objScanningRecord.FileName = Path.GetFileName(strFile)
                objScanningRecord.MachineName = Environment.MachineName
                If blnIsCommonwealthRetry Then
                    objScanningRecord.RetryAttempts = GetRetryAttempts(strFile) + 1
                End If
                strDBStatus = objScanningRecord.AddScanningRecord()
                If strDBStatus <> "OK" Then
                    SendNotificationEmail("Error adding record to database for file Commonwealth: " + strFile, strDBStatus, "ERR15")
                End If
            End If

            If blnUploadToVSI Then
                objScanningRecord = New ScanningRecord
                oServiceLog.WriteLogEntry("Uploading file " + strFile + " to Visionet")
                objScanningRecord.UploadStart = Now.ToString
                strStatusVisi = UploadXMLFile(strTempDir + "\" + Path.GetFileNameWithoutExtension(strFile) + ".xml", strURL)
                objScanningRecord.UploadEnd = Now.ToString
                If strStatusVisi = "OK" Then
                    objScanningRecord.Uploaded = 1
                Else
                    objScanningRecord.Exception = 1
                    objScanningRecord.Message = strStatusVisi
                    oErrorType = GetErrorType(strStatusVisi)
                    Select Case oErrorType
                        Case ErrorType.Retry
                            objScanningRecord.TimeOut = 1
                        Case ErrorType.IncorrectId
                            objScanningRecord.IncorrectId = 1
                        Case ErrorType.InvalidId
                            objScanningRecord.InvalidId = 1
                    End Select
                End If
                objScanningRecord.SystemType = "VSI"
                objScanningRecord.ClientCode = strSplit(1)
                objScanningRecord.LoanNumber = strSplit(3)
                objScanningRecord.OrderNumber = strSplit(2)
                objScanningRecord.FileName = Path.GetFileName(strFile)
                objScanningRecord.MachineName = Environment.MachineName
                If blnIsVSIRetry Then
                    objScanningRecord.RetryAttempts = GetRetryAttempts(strFile) + 1
                End If
                strDBStatus = objScanningRecord.AddScanningRecord()
                If strDBStatus <> "OK" Then
                    SendNotificationEmail("Error adding record to database for file VSI: " + strFile, strDBStatus, "ERR16")
                End If
            End If
            If blnUploadToResware Then
                objScanningRecord = New ScanningRecord
                oServiceLog.WriteLogEntry("Uploading file " + strFile + " to Resware")
                objScanningRecord.UploadStart = Now.ToString
                strStatusResware = UploadReswareFile(strFile, strSplit(3))
                objScanningRecord.UploadEnd = Now.ToString
                If strStatusResware = "OK" Then
                    objScanningRecord.Uploaded = 1
                Else
                    objScanningRecord.Exception = 1
                    objScanningRecord.Message = strStatusResware
                    oErrorType = GetErrorType(strStatusResware)
                    Select Case oErrorType
                        Case ErrorType.Retry
                            objScanningRecord.TimeOut = 1
                        Case ErrorType.IncorrectId
                            objScanningRecord.IncorrectId = 1
                        Case ErrorType.InvalidId
                            objScanningRecord.InvalidId = 1
                    End Select
                End If
                objScanningRecord.SystemType = "Resware"
                objScanningRecord.ClientCode = strSplit(1)
                objScanningRecord.LoanNumber = strSplit(3)
                objScanningRecord.OrderNumber = strSplit(2)
                objScanningRecord.FileName = Path.GetFileName(strFile)
                objScanningRecord.MachineName = Environment.MachineName
                If blnIsReswareRetry Then
                    objScanningRecord.RetryAttempts = GetRetryAttempts(strFile) + 1
                End If
                strDBStatus = objScanningRecord.AddScanningRecord()
                If strDBStatus <> "OK" Then
                    SendNotificationEmail("Error adding record to database for file Resware: " + strFile, strDBStatus, "ERR17")
                End If
            End If
            If blnUploadToSilkResware Then
                objScanningRecord = New ScanningRecord
                oServiceLog.WriteLogEntry("Uploading file " + strFile + " to SilkResware")
                objScanningRecord.UploadStart = Now.ToString
                strStatusSilkResware = UploadSilkReswareFile(strFile, strSplit(3))
                objScanningRecord.UploadEnd = Now.ToString
                If strStatusSilkResware = "OK" Then
                    objScanningRecord.Uploaded = 1
                Else
                    objScanningRecord.Exception = 1
                    objScanningRecord.Message = strStatusSilkResware
                    oErrorType = GetErrorType(strStatusSilkResware)
                    Select Case oErrorType
                        Case ErrorType.Retry
                            objScanningRecord.TimeOut = 1
                        Case ErrorType.IncorrectId
                            objScanningRecord.IncorrectId = 1
                        Case ErrorType.InvalidId
                            objScanningRecord.InvalidId = 1
                    End Select
                End If
                objScanningRecord.SystemType = "SilkResware"
                objScanningRecord.ClientCode = strSplit(1)
                objScanningRecord.LoanNumber = strSplit(3)
                objScanningRecord.OrderNumber = strSplit(2)
                objScanningRecord.FileName = Path.GetFileName(strFile)
                objScanningRecord.MachineName = Environment.MachineName
                If blnIsReswareRetry Then
                    objScanningRecord.RetryAttempts = GetRetryAttempts(strFile) + 1
                End If
                strDBStatus = objScanningRecord.AddScanningRecord()
                If strDBStatus <> "OK" Then
                    SendNotificationEmail("Error adding record to database for file SilkResware: " + strFile, strDBStatus, "ERR18")
                End If
            End If
            intVSIError = 0
            intWFGError = 0
            intReswareError = 0
            intCommonwealthError = 0
            intSilkReswareError = 0
            intAllError = 0
            If strStatusVisi = "OK" AndAlso strStatusWFG = "OK" AndAlso strStatusResware = "OK" AndAlso strStatusCommonwealth = "OK" AndAlso strStatusSilkResware = "OK" Then
                strStatusMessage.Append("OK")
            Else
                strStatusMessage.Append("Error uploading to: " + vbCrLf)
                If strStatusVisi <> "OK" AndAlso strStatusWFG <> "OK" AndAlso strStatusResware <> "OK" AndAlso strStatusCommonwealth <> "OK" Then
                    strStatusMessage.Append("Visionet: " + strStatusVisi + vbCrLf + "WFG: " + strStatusWFG + vbCrLf + "Resware: " + strStatusResware + vbCrLf + "Commonwealth: " + strStatusCommonwealth)
                    intAllError = 1
                Else
                    If strStatusVisi <> "OK" Then
                        strStatusMessage.Append("Visionet: " + strStatusVisi + vbCrLf)
                        intVSIError = 1
                    End If
                    If strStatusWFG <> "OK" Then
                        strStatusMessage.Append("WFG:" + strStatusWFG + vbCrLf)
                        intWFGError = 1
                    End If
                    If strStatusResware <> "OK" Then
                        strStatusMessage.Append("Resware:" + strStatusResware + vbCrLf)
                        intReswareError = 1
                    End If
                    If strStatusCommonwealth <> "OK" Then
                        strStatusMessage.Append("Commonwealth:" + strStatusCommonwealth + vbCrLf)
                        intCommonwealthError = 1
                    End If
                    If strStatusSilkResware <> "OK" Then
                        strStatusMessage.Append("Silk Resware:" + strStatusSilkResware + vbCrLf)
                        intSilkReswareError = 1
                    End If
                End If
            End If
            Return strStatusMessage.ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Private Function UploadXMLFile(strFile As String, strSendToUrl As String) As String
        Dim strStatus As String = "Upload Error"
        Try
            Dim request As HttpWebRequest
            request = WebRequest.Create(strSendToUrl)
            request.Method = "POST"
            request.ContentType = "text/xml"
            request.Timeout = intUploadTimeout
            Dim writer As StreamWriter = New StreamWriter(request.GetRequestStream)
            writer.WriteLine(GetTextFromXMLFile(strFile))
            writer.Close()
            Dim response As HttpWebResponse = request.GetResponse
            Dim dataStream As Stream = response.GetResponseStream
            Dim reader As StreamReader = New StreamReader(dataStream)
            Dim strResponse As String = reader.ReadToEnd
            reader.Close()
            dataStream.Close()
            response.Close()

            oServiceLog.WriteLogEntry("Upload Response = " + strResponse)
            Dim xmlDoc As XmlDocument = New XmlDocument
            Dim xmlNodeValidated As XmlNode
            Dim xmlNodeErrorDesc As XmlNode
            Dim xmlNodeProviderOrderNbr As XmlNode
            xmlDoc.LoadXml(strResponse)
            xmlNodeValidated = xmlDoc.SelectSingleNode("//" + "Validated")
            xmlNodeProviderOrderNbr = xmlDoc.SelectSingleNode("//" + "ProviderOrderNbr")
            If xmlNodeValidated.InnerText = "Yes" And Not IsNothing(xmlNodeProviderOrderNbr) Then
                strStatus = "OK"
            Else
                xmlNodeErrorDesc = xmlDoc.SelectSingleNode("//" + "Description")
                If Not IsNothing(xmlNodeErrorDesc) Then
                    strStatus = xmlNodeErrorDesc.InnerText
                Else
                    strStatus = "Order number does not exist in the AtClose systsm."
                End If
            End If
            Return strStatus

        Catch webEx As WebException
            Return webEx.Message
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Private Function UploadReswareFile(strFile As String, strFileNumber As String) As String
        Dim strStatus = "Error"
        Dim response As LinearReceiveNoteService.ReceiveNoteResponse

        Try
            Dim notesDocument(0) As LinearReceiveNoteService.ReceiveNoteDocument
            notesDocument(0) = New LinearReceiveNoteService.ReceiveNoteDocument
            notesDocument(0).FileName = Path.GetFileName(strFile)

            If (strFile.Contains("Unrecorded")) Then
                notesDocument(0).DocumentTypeID = 1230
                notesDocument(0).Description = "Unrecorded Documents"
            ElseIf (strFile.Contains("Recorded")) Then
                notesDocument(0).DocumentTypeID = 1571
                notesDocument(0).Description = "Recorded Document"
            Else
                notesDocument(0).DocumentTypeID = 1005
                notesDocument(0).Description = "Signed Post Closing Package"
            End If

            oServiceLog.WriteLogEntry("Converting file " + strFile + "to byte array.")
            oServiceLog.WriteLogEntry("   Resware Outbound info:")
            oServiceLog.WriteLogEntry("    Document Classification: DocumentTypeID=" + notesDocument(0).DocumentTypeID.ToString() + " Description = " + notesDocument(0).Description)
            notesDocument(0).DocumentBody = ConvertFileToByteArray(strFile)
            oServiceLog.WriteLogEntry("File " + strFile + "converted")

            Dim notesData As LinearReceiveNoteService.ReceiveNoteData = New LinearReceiveNoteService.ReceiveNoteData
            notesData.FileNumber = strFileNumber
            notesData.Documents = notesDocument
            oServiceLog.WriteLogEntry("Setting up Resware service client and credentials.")
            Dim notesService As LinearReceiveNoteService.ReceiveNoteServiceClient = New LinearReceiveNoteService.ReceiveNoteServiceClient
            oServiceLog.WriteLogEntry("Client created")
            notesService.ClientCredentials.UserName.UserName = System.Configuration.ConfigurationManager.AppSettings("ReswareLoginUser")
            oServiceLog.WriteLogEntry("User name set")
            notesService.ClientCredentials.UserName.Password = System.Configuration.ConfigurationManager.AppSettings("ReswareLoginPassword")
            oServiceLog.WriteLogEntry("Password set")
            oServiceLog.WriteLogEntry("Invoking Resware upload method")
            response = notesService.ReceiveNote(notesData)
            oServiceLog.WriteLogEntry("Resware response received = " + response.ResponseCode.ToString)
            If response.ResponseCode = LinearReceiveNoteService.ReceiveNoteResponseCode.SUCCESS Then
                Return "OK"
            Else
                Return response.Message
            End If

        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Private Function UploadSilkReswareFile(strFile As String, strFileNumber As String) As String
        Dim strStatus = "Error"
        Dim response As SilkReceiveNoteService.ReceiveNoteResponse

        Try
            Dim notesDocument(0) As SilkReceiveNoteService.ReceiveNoteDocument
            notesDocument(0) = New SilkReceiveNoteService.ReceiveNoteDocument
            notesDocument(0).FileName = Path.GetFileName(strFile)
            notesDocument(0).DocumentTypeID = 1177
            notesDocument(0).Description = "Post Closing Package - Not Reviewed"
            oServiceLog.WriteLogEntry("Converting file " + strFile + "to byte array.")
            notesDocument(0).DocumentBody = ConvertFileToByteArray(strFile)
            oServiceLog.WriteLogEntry("File " + strFile + "converted")

            Dim notesData As SilkReceiveNoteService.ReceiveNoteData = New SilkReceiveNoteService.ReceiveNoteData
            notesData.FileNumber = strFileNumber
            notesData.Documents = notesDocument
            oServiceLog.WriteLogEntry("Setting up SilkResware service client and credentials.")
            Dim notesService As SilkReceiveNoteService.ReceiveNoteServiceClient = New SilkReceiveNoteService.ReceiveNoteServiceClient
            oServiceLog.WriteLogEntry("Client created")
            notesService.ClientCredentials.UserName.UserName = System.Configuration.ConfigurationManager.AppSettings("SilkReswareLoginUser")
            oServiceLog.WriteLogEntry("User name set")
            notesService.ClientCredentials.UserName.Password = System.Configuration.ConfigurationManager.AppSettings("SilkReswareLoginPassword")
            oServiceLog.WriteLogEntry("Password set")
            oServiceLog.WriteLogEntry("Invoking SilkResware upload method")
            response = notesService.ReceiveNote(notesData)
            oServiceLog.WriteLogEntry("SilkResware response received = " + response.ResponseCode.ToString)
            If response.ResponseCode = SilkReceiveNoteService.ReceiveNoteResponseCode.SUCCESS Then
                oServiceLog.WriteLogEntry("Resware [notesService.ReceiveNote] service sucesfully called")
                Return "OK"
            Else
                oServiceLog.WriteLogEntry("Resware [notesService.ReceiveNote] FAILURE")
                oServiceLog.WriteLogEntry("Response Failure Code = " & response.ResponseCode)
                Return response.Message
            End If

        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Private Function GetTextFromXMLFile(strFile As String) As String
        Dim reader As StreamReader = New StreamReader(strFile)
        Dim ret As String = reader.ReadToEnd
        reader.Close()
        Return ret
    End Function

    Private Function ConvertFileToBase64(ByVal fileName As String) As String
        Dim ReturnValue As String = ""
        If My.Computer.FileSystem.FileExists(fileName) Then
            Using BinaryFile As FileStream = New FileStream(fileName, FileMode.Open)
                Dim BinRead As BinaryReader = New BinaryReader(BinaryFile)
                Dim BinBytes As Byte() = BinRead.ReadBytes(CInt(BinaryFile.Length))
                ReturnValue = Convert.ToBase64String(BinBytes)
                BinaryFile.Close()
            End Using
        End If
        Return ReturnValue
    End Function

    Private Function ConvertFileToByteArray(fileName As String) As Byte()
        Dim fInfo As New FileInfo(fileName)
        Dim numBytes As Long = fInfo.Length
        Dim fStream As New FileStream(fileName, FileMode.Open, FileAccess.Read)
        Dim br As New BinaryReader(fStream)
        Dim data As Byte() = br.ReadBytes(CInt(numBytes))
        br.Close()
        fStream.Close()
        Return data
    End Function

    Public Sub SendNotificationEmail(strFileName As String, strError As String, strERRID As String)
        Dim mailMessage As New MailMessage
        Dim strSplit() As String = strMailToUser1.Split(";")
        Try
            mailMessage.From = New MailAddress(strMailFromAddress)
            For Each strAddress As String In strSplit
                mailMessage.To.Add(strAddress)
            Next
            If strMailToUser2 <> String.Empty Then
                mailMessage.To.Add(strMailToUser2)
            End If
            mailMessage.Subject = "AtCloseUploadService exception: " + strFileName
            mailMessage.Body = "There was a problem processing and uploading the file: " + strFileName + vbCrLf + "on " + Environment.MachineName + vbCrLf + "Error: " + strError + vbCrLf + vbCrLf + vbCrLf
            mailMessage.Body = mailMessage.Body + "ERRID=[" + strERRID + "]"
            mailMessage.IsBodyHtml = True
            Dim SMTP As New SmtpClient("smtp.office365.com")
            SMTP.UseDefaultCredentials = False
            SMTP.Credentials = New System.Net.NetworkCredential(strMailFromUser, strMailFromPassword)
            SMTP.Port = 587
            SMTP.EnableSsl = True
            SMTP.DeliveryMethod = SmtpDeliveryMethod.Network
            SMTP.Timeout = 20000
            'SMTP.TargetName = "STARTTLS/smtp.office365.com"
            SMTP.Send(mailMessage)
        Catch ex As Exception
            oServiceLog.WriteLogEntry("Error sending notification email: " + ex.Message)
        End Try
    End Sub

    Private Function GetErrorType(strStatusMessage As String) As ErrorType
        If strStatusMessage.Contains("underlying connection was closed") Then
            Return ErrorType.Retry
        ElseIf strStatusMessage.Contains("request channel timed out") Then
            Return ErrorType.Retry
        ElseIf strStatusMessage.Contains("Given combination of Transaction ID and Unique ID") Then
            Return ErrorType.InvalidId
        ElseIf strStatusMessage.Contains("Invalid FileNumber") Then
            Return ErrorType.InvalidId
        ElseIf strStatusMessage.Contains("Order number does not exist") Then
            Return ErrorType.InvalidId
        ElseIf strStatusMessage.Contains("Gateway Timeout") Then
            Return ErrorType.Retry
        ElseIf strStatusMessage.Contains("Session has unexpectedly closed") Then
            Return ErrorType.Retry
        ElseIf strStatusMessage.Contains("Server Unavailable") Then
            Return ErrorType.Retry
        Else
            Return ErrorType.System
        End If
    End Function

    Private Function GetRetryAttempts(strFileName As String) As Integer
        Dim connSQL As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("ConnectionStringLocal"))
        Dim cmdGetRetryAttempts As SqlCommand = New SqlCommand
        Dim intRetryAttempts As Integer = 0
        Try
            connSQL.Open()
            cmdGetRetryAttempts.CommandText = "SELECT MAX(RetryAttemps) FROM tblScanning WHERE FileName = '" + Path.GetFileName(strFileName) + "' GROUP BY FileName"
            intRetryAttempts = cmdGetRetryAttempts.ExecuteScalar
        Catch
        End Try
        Return intRetryAttempts
    End Function

End Class
