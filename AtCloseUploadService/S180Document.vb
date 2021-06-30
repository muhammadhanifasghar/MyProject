Public Class S180Document
    Private m_Description As String
    Private m_DocDate As String
    Private m_DocVersion As String
    Private m_DocStatus As String
    Private m_DocName As String
    Private m_FileName As String
    Private m_FileType As String
    Private m_ExpireDate As String
    Private m_Content As String
    Private m_EncodeType As String
    Private m_DocType As String
    Private m_DocTypeOtherDesc As String

    Public Property Description As String
        Get
            Return m_Description
        End Get
        Set(value As String)
            m_Description = value
        End Set
    End Property

    Public Property DocDate As String
        Get
            Return m_DocDate
        End Get
        Set(value As String)
            m_DocDate = value
        End Set
    End Property

    Public Property DocVersion As String
        Get
            Return m_DocVersion
        End Get
        Set(value As String)
            m_DocVersion = value
        End Set
    End Property

    Public Property DocStatus As String
        Get
            Return m_DocStatus
        End Get
        Set(value As String)
            m_DocStatus = value
        End Set
    End Property

    Public Property DocName As String
        Get
            Return m_DocName
        End Get
        Set(value As String)
            m_DocName = value
        End Set
    End Property

    Public Property FileName As String
        Get
            Return m_FileName
        End Get
        Set(value As String)
            m_FileName = value
        End Set
    End Property

    Public Property FileType As String
        Get
            Return m_FileType
        End Get
        Set(value As String)
            m_FileType = value
        End Set
    End Property

    Public Property ExpireDate As String
        Get
            Return m_ExpireDate
        End Get
        Set(value As String)
            m_ExpireDate = value
        End Set
    End Property

    Public Property Content As String
        Get
            Return m_Content
        End Get
        Set(value As String)
            m_Content = value
        End Set
    End Property

    Public Property EncodeType As String
        Get
            Return m_EncodeType
        End Get
        Set(value As String)
            m_EncodeType = value
        End Set
    End Property

    Public Property DocType As String
        Get
            Return m_DocType
        End Get
        Set(value As String)
            m_DocType = value
        End Set
    End Property

    Public Property DocTypeOtherDesc As String
        Get
            Return m_DocTypeOtherDesc
        End Get
        Set(value As String)
            m_DocTypeOtherDesc = value
        End Set
    End Property
End Class
