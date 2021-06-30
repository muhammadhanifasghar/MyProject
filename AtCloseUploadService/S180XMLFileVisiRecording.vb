Imports System.Xml

Public Class S180XMLFileVisiRecording
    Private m_ProductList As List(Of S180Product)
    Private m_TransactionId As String
    Private m_ProviderOrderNbr As String

    Public Property ProductList As List(Of S180Product)
        Get
            Return m_ProductList
        End Get
        Set(value As List(Of S180Product))
            m_ProductList = value
        End Set
    End Property

    Public Property TransactionId As String
        Get
            Return m_TransactionId
        End Get
        Set(value As String)
            m_TransactionId = value
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

    Public Function WriteToFile(strFullFileName As String) As String
        Dim doc As XmlDocument = New XmlDocument
        Dim product As XmlElement
        Dim productProviderOrderNbr As XmlElement
        Dim documentList As XmlElement
        Dim documentCount As XmlElement
        Dim document As XmlElement
        Dim documentDescription As XmlElement
        Dim documentContent As XmlElement

        Try
            Dim transaction As XmlElement = doc.CreateElement("Transaction")
            Dim transactionId As XmlElement = doc.CreateElement("TransactionId")
            transactionId.InnerText = Me.TransactionId
            transaction.AppendChild(transactionId)
            Dim productList As XmlElement = doc.CreateElement("ProductList")
            Dim productListCount As XmlElement = doc.CreateElement("Count")
            productListCount.InnerText = m_ProductList.Count
            productList.AppendChild(productListCount)
            For Each m_Product As S180Product In m_ProductList
                product = doc.CreateElement("Product")
                productProviderOrderNbr = doc.CreateElement("ProviderOrderNbr")
                productProviderOrderNbr.InnerText = Me.ProviderOrderNbr
                product.AppendChild(productProviderOrderNbr)
                documentList = doc.CreateElement("DocumentList")
                product.AppendChild(documentList)
                documentCount = doc.CreateElement("Count")
                documentCount.InnerText = m_Product.DocumentList.Count
                documentList.AppendChild(documentCount)
                For Each m_Document As S180Document In m_Product.DocumentList
                    document = doc.CreateElement("Document")
                    documentDescription = doc.CreateElement("Description")
                    documentDescription.InnerText = m_Document.Description
                    documentContent = doc.CreateElement("Content")
                    documentContent.InnerText = m_Document.Content
                    document.AppendChild(documentDescription)
                    document.AppendChild(documentContent)
                    documentList.AppendChild(document)
                Next
                productList.AppendChild(product)
            Next
            transaction.AppendChild(productList)
            doc.AppendChild(transaction)

            doc.Save(strFullFileName)
            oServiceLog.WriteLogEntry("XML file " + strFullFileName + "created")
            Return "OK"
        Catch ex As Exception
            oServiceLog.WriteLogEntry("Error creating xml file: " + ex.Message)
            Return "Error"
        End Try
    End Function
End Class
