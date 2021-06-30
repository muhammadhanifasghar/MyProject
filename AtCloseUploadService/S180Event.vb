Public Class S180Event
    Private m_Code As String
    Private m_EventDate As String
    Private m_OccurDate As String
    Private m_ClientCode As String
    Private m_ClientDesc As String
    Private m_Comment As String
    Private m_EventId As String

    Public Property Code As String
        Get
            Return m_Code
        End Get
        Set(value As String)
            m_Code = value
        End Set
    End Property

    Public Property EventDate As String
        Get
            Return m_EventDate
        End Get
        Set(value As String)
            m_EventDate = value
        End Set
    End Property

    Public Property OccurDate As String
        Get
            Return m_OccurDate
        End Get
        Set(value As String)
            m_OccurDate = value
        End Set
    End Property

    Public Property ClientCode As String
        Get
            Return m_ClientCode
        End Get
        Set(value As String)
            m_ClientCode = value
        End Set
    End Property

    Public Property ClientDesc As String
        Get
            Return m_ClientDesc
        End Get
        Set(value As String)
            m_ClientDesc = value
        End Set
    End Property

    Public Property Comment As String
        Get
            Return m_Comment
        End Get
        Set(value As String)
            m_Comment = value
        End Set
    End Property

    Public Property EventId As String
        Get
            Return m_EventId
        End Get
        Set(value As String)
            m_EventId = value
        End Set
    End Property
End Class
