Imports System.IO
Imports System.Text

#If False Then

Date        Author              Comments
4-08-2016   Norman Gottschalk   Added a check to see if the file is open prior to processing to remove the error messages
                                "because it is being used by another person"
05-09-2016  Norman Gottschalk   For Linear (Resware) changed the document type to be Unrecorded(1230), Recorded Documents-Other (1571) 
                                and the original (default) Signed Post Closing Package (1005)

#End If

Public Class AtCloseUploadService

    Protected Overrides Sub OnStart(ByVal args() As String)
        InitProcessing()
        Dim timer As System.Timers.Timer = New System.Timers.Timer
        timer.Interval = intPollingIntervalMins * 60000
        AddHandler timer.Elapsed, AddressOf Me.OnTimer
        timer.Start()
    End Sub

    Protected Overrides Sub OnStop()
        If oServiceLog.LogFlag = 1 Then
            oServiceLog.CloseLogFile()
        End If
    End Sub

    Private Sub OnTimer(sender As Object, e As Timers.ElapsedEventArgs)
        StartProcessing()
    End Sub

    Function IsFileOpen(file As FileInfo)
        Dim stream As FileStream = Nothing
        Try
            stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None)
            stream.Close()
            IsFileOpen = False
        Catch ex As Exception
            IsFileOpen = True
        End Try
    End Function

    Private Sub InitProcessing()
        Dim strLogFlag As String = "0"
        strLogFlag = System.Configuration.ConfigurationManager.AppSettings("logFlag")
        strLogDir = System.Configuration.ConfigurationManager.AppSettings("logDir")
        If Not Directory.Exists(strLogDir) Then
            Directory.CreateDirectory(strLogDir)
        End If
        oServiceLog = New ServiceLog
        If strLogFlag = "1" Then
            oServiceLog.LogFlag = 1
            oServiceLog.OpenLogFile()
        Else
            oServiceLog.LogFlag = 0
        End If
        oServiceLog.WriteLogEntry("Log Dir = " + strLogDir)
        oServiceLog.WriteLogEntry("Starting service 05-29-2020 v f")
        strURL = System.Configuration.ConfigurationManager.AppSettings("url")
        oServiceLog.WriteLogEntry("URL = " + strURL)
        strTimiosURL = System.Configuration.ConfigurationManager.AppSettings("timiosurl")
        oServiceLog.WriteLogEntry("TimiosURL = " + strTimiosURL)
        strVisiRecordingURL = System.Configuration.ConfigurationManager.AppSettings("visirecordingurl")
        oServiceLog.WriteLogEntry("VisiRecordingURL = " + strVisiRecordingURL)
        strCommonwealthURL = System.Configuration.ConfigurationManager.AppSettings("commonwealthurl")
        oServiceLog.WriteLogEntry("CommonWealthURL = " + strCommonwealthURL)
        strListenPath = System.Configuration.ConfigurationManager.AppSettings("listenPath")
        oServiceLog.WriteLogEntry("Listen Path = " + strListenPath)
        strProcessedPath = System.Configuration.ConfigurationManager.AppSettings("processedPath")
        oServiceLog.WriteLogEntry("Processed Path = " + strProcessedPath)
        If Not Directory.Exists(strProcessedPath) Then
            Directory.CreateDirectory(strProcessedPath)
        End If
        strExceptionPath = System.Configuration.ConfigurationManager.AppSettings("exceptionPath")
        oServiceLog.WriteLogEntry("Exception Path = " + strExceptionPath)
        If Not Directory.Exists(strExceptionPath) Then
            Directory.CreateDirectory(strExceptionPath)
        End If
        intPollingIntervalMins = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings("pollingIntervalMins"))
        oServiceLog.WriteLogEntry("Polling Interval = " + intPollingIntervalMins.ToString + " mins")
        strTempDir = System.Configuration.ConfigurationManager.AppSettings("tempDir")
        oServiceLog.WriteLogEntry("Temp Dir = " + strTempDir)
        If Not Directory.Exists(strTempDir) Then
            Directory.CreateDirectory(strTempDir)
        End If

        strMailFromAddress = System.Configuration.ConfigurationManager.AppSettings("MailFromAddress")
        oServiceLog.WriteLogEntry("MailFromAddress = " + strMailFromAddress)
        strMailFromUser = System.Configuration.ConfigurationManager.AppSettings("MailFromUser")
        oServiceLog.WriteLogEntry("MailFromUser = " + strMailFromUser)
        strMailFromPassword = System.Configuration.ConfigurationManager.AppSettings("MailFromPassword")
        oServiceLog.WriteLogEntry("MailFromPassword = " + strMailFromPassword)
        strMailToUser1 = System.Configuration.ConfigurationManager.AppSettings("MailToUser1")
        oServiceLog.WriteLogEntry("MailToUser1 = " + strMailToUser1)
        strMailToUser2 = System.Configuration.ConfigurationManager.AppSettings("MailToUser2")
        oServiceLog.WriteLogEntry("MailToUser2 = " + strMailToUser2)
        intUploadTimeout = System.Configuration.ConfigurationManager.AppSettings("UploadTimeout")
        intVisiRecordingTimeout = System.Configuration.ConfigurationManager.AppSettings("UploadTimeout")

    End Sub

    Private Sub StartProcessing()
        If intProcessing = 0 Then
            intProcessing = 1
            Dim dirInfo As DirectoryInfo
            Dim fiFiles() As FileInfo
            Dim strStatus As String = "Error"
            Dim strErrorPrefix As StringBuilder = New StringBuilder
            Dim oFileProcessor As FileProcessor = New FileProcessor
            Try
                oServiceLog.WriteLogEntry("Checking for files in: " + strListenPath)
                dirInfo = New IO.DirectoryInfo(strListenPath)
                fiFiles = dirInfo.GetFiles().OrderByDescending(Function(fi) fi.LastWriteTime).ToArray()
                If fiFiles.Length > 0 Then
                    oServiceLog.WriteLogEntry(fiFiles.Length.ToString + " files found")
                    For Each fiFile As FileInfo In fiFiles
                        oServiceLog.WriteLogEntry("Begin processing for file:" + fiFile.Name)

                        If fiFile.FullName.Contains("Transaction") Then
                            Try
                                File.Move(fiFile.FullName, strProcessedPath + "\" + Path.GetFileName(fiFile.FullName))
                            Catch ex As Exception
                                If File.Exists(strProcessedPath + "\" + Path.GetFileName(fiFile.FullName)) Then
                                    File.Delete(fiFile.FullName)
                                End If
                                oServiceLog.WriteLogEntry("Error moving transaction file: " + fiFile.FullName)
                                oFileProcessor.SendNotificationEmail(fiFile.FullName, "Error moving transaction file:", "ERR2")
                            End Try
                        Else
                            strStatus = oFileProcessor.ProcessFile(fiFile.FullName)
                            If strStatus = "OK" Then
                                Try
                                    File.Move(fiFile.FullName, strProcessedPath + "\" + Path.GetFileName(fiFile.FullName))
                                Catch
                                    If File.Exists(strProcessedPath + "\" + Path.GetFileName(fiFile.FullName)) Then
                                        File.Delete(fiFile.FullName)
                                    End If
                                    oServiceLog.WriteLogEntry("Error trying to move file to processing folder: " + fiFile.FullName)
                                    oFileProcessor.SendNotificationEmail(fiFile.FullName, "Error trying to move file to processing folder:", "ERR1")
                                End Try
                                File.Delete(strTempDir + "\" + Path.GetFileNameWithoutExtension(fiFile.FullName) + ".xml")
                            Else
                                strErrorPrefix.Clear()
                                If intAllError = 1 Then
                                    strErrorPrefix.Append("11111111")
                                    Try
                                        File.Move(fiFile.FullName, strExceptionPath + "\" + strErrorPrefix.ToString + Path.GetFileName(fiFile.FullName))
                                    Catch ex As Exception
                                        If File.Exists(strExceptionPath(+"\" + strErrorPrefix.ToString + Path.GetFileName(fiFile.FullName))) Then
                                            File.Delete(fiFile.FullName)
                                        End If
                                        oServiceLog.WriteLogEntry("Error trying to move file to exception folder: " + fiFile.FullName)
                                        oFileProcessor.SendNotificationEmail(fiFile.FullName, "Error trying to move file to exception folder:", "ERR2")
                                    End Try
                                    File.Delete(strTempDir + "\" + Path.GetFileNameWithoutExtension(fiFile.FullName) + ".xml")
                                    oFileProcessor.SendNotificationEmail(fiFile.FullName, strStatus, "ERR3")
                                Else
                                    If intVSIError Then
                                        strErrorPrefix.Append("1")
                                    Else
                                        strErrorPrefix.Append("0")
                                    End If
                                    If intWFGError Then
                                        strErrorPrefix.Append("1")
                                    Else
                                        strErrorPrefix.Append("0")
                                    End If
                                    If intReswareError Then
                                        strErrorPrefix.Append("1")
                                    Else
                                        strErrorPrefix.Append("0")
                                    End If
                                    If intCommonwealthError Then
                                        strErrorPrefix.Append("1")
                                    Else
                                        strErrorPrefix.Append("0")
                                    End If
                                    If intSilkReswareError Then
                                        strErrorPrefix.Append("1")
                                    Else
                                        strErrorPrefix.Append("0")
                                    End If
                                    If intHowardHannaError Then
                                        strErrorPrefix.Append("1")
                                    Else
                                        strErrorPrefix.Append("0")
                                    End If
                                    If intVLRError Then
                                        strErrorPrefix.Append("1")
                                    Else
                                        strErrorPrefix.Append("0")
                                    End If
                                    If intTimiosError Then
                                        strErrorPrefix.Append("1")
                                    Else
                                        strErrorPrefix.Append("0")
                                    End If
                                    If intVisiRecordingError Then
                                        strErrorPrefix.Append("1")
                                    Else
                                        strErrorPrefix.Append("0")
                                    End If
                                End If
                                Try
                                    File.Move(fiFile.FullName, strExceptionPath + "\" + strErrorPrefix.ToString + Path.GetFileName(fiFile.FullName))
                                Catch
                                    If File.Exists(strExceptionPath + "\" + strErrorPrefix.ToString + Path.GetFileName(fiFile.FullName)) Then
                                        File.Delete(fiFile.FullName)
                                    End If
                                    oServiceLog.WriteLogEntry("Error trying to move file to exception folder: " + fiFile.FullName)
                                    oFileProcessor.SendNotificationEmail(fiFile.FullName, "Error trying to move file to exception folder:", "ERR4")
                                End Try
                                File.Delete(strTempDir + "\" + Path.GetFileNameWithoutExtension(fiFile.FullName) + ".xml")
                                oFileProcessor.SendNotificationEmail(fiFile.FullName, strStatus, "ERR5")
                            End If
                        End If
                        oServiceLog.WriteLogEntry("End processing for file:" + fiFile.FullName)
                    Next
                Else
                    oServiceLog.WriteLogEntry("No files found")
                End If
            Catch ex As Exception
                oServiceLog.WriteLogEntry("Error processing files: " + ex.Message)
                oFileProcessor.SendNotificationEmail("System Error", ex.Message, "ERR6")
            End Try
            intProcessing = 0
        End If
    End Sub

    Public Sub TestService()
        InitProcessing()
        StartProcessing()
    End Sub
End Class
