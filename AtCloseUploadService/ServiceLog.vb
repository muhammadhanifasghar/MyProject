Imports System.IO

Public Class ServiceLog
    Private m_LogFlag As Integer = 0
    Private strCurrentLog As String
    Private oLogFile As StreamWriter

    Public Property LogFlag As Integer
        Get
            Return m_LogFlag
        End Get
        Set(value As Integer)
            m_LogFlag = value
        End Set
    End Property

    Public Sub OpenLogFile()
        Dim strCurrentLog As String
        strCurrentLog = strLogDir + "\AtCloseUploadServiceLog_" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".log"
        oLogFile = New StreamWriter(strCurrentLog)
    End Sub

    Public Sub CloseLogFile()
        oLogFile.Close()
        oLogFile.Dispose()
    End Sub

    Public Sub WriteLogEntry(strMessage As String)
        If Me.LogFlag = 1 Then
            oLogFile.WriteLine(String.Format(Date.Now, "MM/dd/yyyy HH:mm:ss") + " " + strMessage)
            oLogFile.Flush()
        End If
    End Sub
End Class
