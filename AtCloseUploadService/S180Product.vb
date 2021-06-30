Public Class S180Product
    Private m_UniqueId As String
    Private m_Instructions As String
    Private m_Reference As String
    Private m_ProviderOrderNbr As String
    Private m_LoanNumber As String
    Private m_LoanType As String
    Private m_Event As S180Event
    Private m_DocumentList As List(Of S180Document)

    Public Property UniqueId As String
        Get
            Return m_UniqueId
        End Get
        Set(value As String)
            m_UniqueId = value
        End Set
    End Property

    Public Property Instructions As String
        Get
            Return m_Instructions
        End Get
        Set(value As String)
            m_Instructions = value
        End Set
    End Property

    Public Property Reference As String
        Get
            Return m_Reference
        End Get
        Set(value As String)
            m_Reference = value
        End Set
    End Property

    Public Property ProviderOrderNbr As String
        Get
            Return m_ProviderOrderNbr
        End Get
        Set(value As String)
            m_ProviderOrderNbr = value
        End Set
    End Property
    Public Property LoanNumber As String
        Get
            Return m_LoanNumber
        End Get
        Set(value As String)
            m_LoanNumber = value
        End Set
    End Property
    Public Property LoanType As String
        Get
            Return m_LoanType
        End Get
        Set(value As String)
            m_LoanType = value
        End Set
    End Property

    Public Property S180Event As S180Event
        Get
            Return m_Event
        End Get
        Set(value As S180Event)
            m_Event = value
        End Set
    End Property

    Public Property DocumentList As List(Of S180Document)
        Get
            Return m_DocumentList
        End Get
        Set(value As List(Of S180Document))
            m_DocumentList = value
        End Set
    End Property

End Class
