Imports System.Data.SqlClient
Imports System.Text

Public Class ScanningRecord
    Private m_SystemType As String

    Public Property SystemType As String
        Get
            Return m_SystemType
        End Get
        Set(value As String)
            m_SystemType = value
        End Set
    End Property

    Private m_FileName As String

    Public Property FileName As String
        Get
            Return m_FileName
        End Get
        Set(value As String)
            m_FileName = value
        End Set
    End Property

    Private m_OrderNumber As String

    Public Property OrderNumber As String
        Get
            Return m_OrderNumber
        End Get
        Set(value As String)
            m_OrderNumber = value
        End Set
    End Property

    Private m_LoanNumber As String

    Public Property LoanNumber As String
        Get
            Return m_LoanNumber
        End Get
        Set(value As String)
            m_LoanNumber = value
        End Set
    End Property

    Private m_ClientCode As String

    Public Property ClientCode As String
        Get
            Return m_ClientCode
        End Get
        Set(value As String)
            m_ClientCode = value
        End Set
    End Property

    Private m_Uploaded As Integer = 0

    Public Property Uploaded As Integer
        Get
            Return m_Uploaded
        End Get
        Set(value As Integer)
            m_Uploaded = value
        End Set
    End Property

    Private m_Exception As Integer = 0

    Public Property Exception As Integer
        Get
            Return m_Exception
        End Get
        Set(value As Integer)
            m_Exception = value
        End Set
    End Property

    Private m_RetryAttempts As Integer = 0

    Public Property RetryAttempts As Integer
        Get
            Return m_RetryAttempts
        End Get
        Set(value As Integer)
            m_RetryAttempts = value
        End Set
    End Property

    Private m_TimeOut As Integer = 0

    Public Property TimeOut As Integer
        Get
            Return m_TimeOut
        End Get
        Set(value As Integer)
            m_TimeOut = value
        End Set
    End Property

    Private m_InvalidId As Integer = 0

    Public Property InvalidId As Integer
        Get
            Return m_InvalidId
        End Get
        Set(value As Integer)
            m_InvalidId = value
        End Set
    End Property

    Private m_IncorrectId As Integer = 0

    Public Property IncorrectId As Integer
        Get
            Return m_IncorrectId
        End Get
        Set(value As Integer)
            m_IncorrectId = value
        End Set
    End Property

    Private m_InvalidLoanNumber As String
    Public Property InvalidLoanNumber As Integer
        Get
            Return m_InvalidLoanNumber
        End Get
        Set(value As Integer)
            m_InvalidLoanNumber = value
        End Set
    End Property

    Private m_Message As String

    Public Property Message As String
        Get
            Return m_Message
        End Get
        Set(value As String)
            m_Message = value
        End Set
    End Property

    Private m_Created As Date

    Public Property Created As Date
        Get
            Return m_Created
        End Get
        Set(value As Date)
            m_Created = value
        End Set
    End Property

    Private m_Updated As Date

    Public Property Updated As Date
        Get
            Return m_Updated
        End Get
        Set(value As Date)
            m_Updated = value
        End Set
    End Property

    Private m_UploadStart As Date

    Public Property UploadStart As Date
        Get
            Return m_UploadStart
        End Get
        Set(value As Date)
            m_UploadStart = value
        End Set
    End Property

    Private m_UploadEnd As Date

    Public Property UploadEnd As Date
        Get
            Return m_UploadEnd
        End Get
        Set(value As Date)
            m_UploadEnd = value
        End Set
    End Property

    Private m_MachineName As String

    Public Property MachineName As String
        Get
            Return m_MachineName
        End Get
        Set(value As String)
            m_MachineName = value
        End Set
    End Property

    Public Function AddScanningRecord() As String
        Dim connSQL As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("ConnectionStringLocal"))
        Dim cmdAddScanningRecord As SqlCommand = New SqlCommand
        Dim strSQL As StringBuilder = New StringBuilder

        Try
            cmdAddScanningRecord.Connection = connSQL

            strSQL.Append("INSERT INTO tblScanning (ScanningRecordId,")
            strSQL.Append("SystemType,")
            strSQL.Append("FileName,")
            strSQL.Append("OrderNumber,")
            strSQL.Append("LoanNumber,")
            strSQL.Append("ClientCode,")
            strSQL.Append("Uploaded,")
            strSQL.Append("Exception,")
            strSQL.Append("RetryAttempts,")
            strSQL.Append("TimeOut,")
            strSQL.Append("InvalidId,")
            strSQL.Append("IncorrectId,")
            strSQL.Append("Error,")
            strSQL.Append("Created,")
            strSQL.Append("Updated,")
            strSQL.Append("UploadStart,")
            strSQL.Append("UploadEnd,")
            strSQL.Append("Cleared, MachineName) VALUES (")
            strSQL.Append("NEWID(),")
            strSQL.Append(ConvertStringToSQLFieldSyntax(Me.SystemType))
            strSQL.Append(ConvertStringToSQLFieldSyntax(StripErrorCodes(Me.FileName)))
            strSQL.Append(ConvertStringToSQLFieldSyntax(Me.OrderNumber))
            strSQL.Append(ConvertStringToSQLFieldSyntax(Me.LoanNumber))
            strSQL.Append(ConvertStringToSQLFieldSyntax(Me.ClientCode))
            strSQL.Append(Me.Uploaded.ToString + ",")
            strSQL.Append(Me.Exception.ToString + ",")
            strSQL.Append(Me.RetryAttempts.ToString + ",")
            strSQL.Append(Me.TimeOut.ToString + ",")
            strSQL.Append(Me.InvalidId.ToString + ",")
            strSQL.Append(Me.IncorrectId.ToString + ",")
            strSQL.Append(ConvertStringToSQLFieldSyntax(Me.Message))
            strSQL.Append("GETDATE(),")
            strSQL.Append("GETDATE(),")
            strSQL.Append(ConvertStringToSQLFieldSyntax(Me.UploadStart))
            strSQL.Append(ConvertStringToSQLFieldSyntax(Me.UploadEnd))
            strSQL.Append("0,")
            strSQL.Append("'" + Me.MachineName + "')")
            connSQL.Open()
            cmdAddScanningRecord.CommandText = strSQL.ToString
            cmdAddScanningRecord.ExecuteNonQuery()
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    Private Function ConvertStringToSQLFieldSyntax(strField As String) As String
        Return "'" + strField + "',"
    End Function

    Public Function UpdateScanningRecord(strFile As String) As String
        Dim connSQL As SqlConnection = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("ConnectionStringLocal"))
        Dim cmdUpdateRecord As SqlCommand = New SqlCommand
        Dim strSQL As StringBuilder = New StringBuilder
        cmdUpdateRecord.Connection = connSQL

        Try
            strSQL.Append("UPDATE tblScanning SET ")
            strSQL.Append("Uploaded = " + Me.Uploaded.ToString + ",")
            strSQL.Append("Exception = " + Me.Exception.ToString + ",")
            strSQL.Append("Error = " + ConvertStringToSQLFieldSyntax(Me.Message) + ",")
            strSQL.Append("RetryAttempts = " + Me.RetryAttempts.ToString + ",")
            strSQL.Append("TimeOut = " + Me.TimeOut.ToString + ",")
            strSQL.Append("InvalidId = " + Me.InvalidId.ToString + ",")
            strSQL.Append("IncorrectId = " + Me.IncorrectId.ToString + ",")
            strSQL.Append("Updated = GETDATE(),")
            strSQL.Append("UploadStart = " + ConvertStringToSQLFieldSyntax(Me.UploadStart.ToString) + ",")
            strSQL.Append("UploadEnd = " + ConvertStringToSQLFieldSyntax(Me.UploadEnd.ToString) + ",")
            strSQL.Append(" WHERE FileName = " + ConvertStringToSQLFieldSyntax(Me.FileName))
            cmdUpdateRecord.CommandText = strSQL.ToString
            cmdUpdateRecord.ExecuteNonQuery()
            Return "OK"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Private Function StripErrorCodes(strFile As String) As String
        If strFile.IndexOf("P") >= 0 Then
            Return strFile.Substring(strFile.IndexOf("P"), strFile.Length - strFile.IndexOf("P"))
        ElseIf strFile.IndexOf("R") >= 0 Then
            Return strFile.Substring(strFile.IndexOf("R"), strFile.Length - strFile.IndexOf("R"))
        End If
        Return Nothing
    End Function

End Class
