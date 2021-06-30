Imports System.Xml

Public Class S180XMLFile
    Private m_ProductList As List(Of S180Product)
    Private m_TransactionId As String

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

    Public Function WriteToFile(strFullFileName As String, Optional strClientCode As String = "") As String
        Dim doc As XmlDocument = New XmlDocument
        Dim product As XmlElement
        Dim productUniqueId As XmlElement
        Dim productProviderOrderNbr As XmlElement
        Dim loanNumber As XmlElement
        Dim loanType As XmlElement
        Dim m_Event As XmlElement
        Dim eventCode As XmlElement
        Dim eventDate As XmlElement
        Dim eventComment As XmlElement
        Dim eventId As XmlElement
        Dim documentList As XmlElement
        Dim documentCount As XmlElement
        Dim document As XmlElement
        Dim documentDescription As XmlElement
        Dim documentDate As XmlElement
        Dim documentVersion As XmlElement
        Dim documentStatus As XmlElement
        Dim documentName As XmlElement
        Dim documentFileName As XmlElement
        Dim documentFileType As XmlElement
        Dim documentContent As XmlElement
        Dim documentEncodeType As XmlElement
        Dim documentDocType As XmlElement

        Try
            Dim root As XmlElement = doc.CreateElement("STAT")
            doc.AppendChild(root)
            Dim header As XmlElement = doc.CreateElement("Header")
            Dim headerFormat As XmlElement = doc.CreateElement("Format")
            headerFormat.InnerText = "STAT"
            Dim headerAckRef As XmlElement = doc.CreateElement("AckRef")
            headerAckRef.InnerText = "1"
            Dim headerCreated As XmlElement = doc.CreateElement("Created")
            headerCreated.InnerText = Strings.Format(Date.Now, "yyyy/MM/dd HH:mm:ss")
            Dim headerCreatedBy As XmlElement = doc.CreateElement("CreatedBy")
            headerCreatedBy.InnerText = "13392292"
            Dim headerUserName As XmlElement = doc.CreateElement("UserName")
            headerUserName.InnerText = "VSIPC"
            Dim headerSourceApp As XmlElement = doc.CreateElement("SourceApp")
            headerSourceApp.InnerText = "VSIPC"
            Dim headerSourceID As XmlElement = doc.CreateElement("SourceID")
            headerSourceID.InnerText = "1001"
            Dim headerSourceVer As XmlElement = doc.CreateElement("SourceVer")
            headerSourceVer.InnerText = "1"
            Dim headerGMTOffset As XmlElement = doc.CreateElement("GMTOffset")
            headerGMTOffset.InnerText = "6"
            Dim headerVersion As XmlElement = doc.CreateElement("Version")
            headerVersion.InnerText = "1"
            header.AppendChild(headerFormat)
            header.AppendChild(headerAckRef)
            header.AppendChild(headerCreated)
            header.AppendChild(headerCreatedBy)
            header.AppendChild(headerUserName)
            header.AppendChild(headerSourceApp)
            header.AppendChild(headerSourceID)
            header.AppendChild(headerSourceVer)
            header.AppendChild(headerGMTOffset)
            header.AppendChild(headerVersion)
            root.AppendChild(header)
            Dim transaction As XmlElement = doc.CreateElement("Transaction")
            root.AppendChild(transaction)
            Dim transactionId As XmlElement = doc.CreateElement("TransactionId")
            transactionId.InnerText = Me.TransactionId
            transaction.AppendChild(transactionId)
            Dim productList As XmlElement = doc.CreateElement("ProductList")
            Dim productListCount As XmlElement = doc.CreateElement("Count")
            productListCount.InnerText = m_ProductList.Count
            productList.AppendChild(productListCount)
            For Each m_Product As S180Product In m_ProductList
                product = doc.CreateElement("Product")
                productUniqueId = doc.CreateElement("UniqueId")
                productUniqueId.InnerText = m_Product.UniqueId
                productProviderOrderNbr = doc.CreateElement("ProviderOrderNbr")
                productProviderOrderNbr.InnerText = m_Product.ProviderOrderNbr
                product.AppendChild(productUniqueId)
                product.AppendChild(productProviderOrderNbr)
                If (strClientCode <> "" And strClientCode.Equals("3858")) Then
                    loanNumber = doc.CreateElement("LoanNumber")
                    loanNumber.InnerText = m_Product.LoanNumber
                    product.AppendChild(loanNumber)
                    loanType = doc.CreateElement("LoanType")
                    loanType.InnerText = m_Product.LoanType
                    product.AppendChild(loanType)
                End If
                m_Event = doc.CreateElement("Event")
                eventCode = doc.CreateElement("Code")
                eventCode.InnerText = m_Product.S180Event.Code
                eventDate = doc.CreateElement("EventDate")
                eventDate.InnerText = m_Product.S180Event.EventDate
                eventComment = doc.CreateElement("Comment")
                eventComment.InnerText = m_Product.S180Event.Comment
                eventId = doc.CreateElement("EventId")
                eventId.InnerText = m_Product.S180Event.EventId
                m_Event.AppendChild(eventCode)
                m_Event.AppendChild(eventDate)
                m_Event.AppendChild(eventComment)
                m_Event.AppendChild(eventId)
                product.AppendChild(m_Event)
                documentList = doc.CreateElement("DocumentList")
                product.AppendChild(documentList)
                documentCount = doc.CreateElement("Count")
                documentCount.InnerText = m_Product.DocumentList.Count
                documentList.AppendChild(documentCount)
                For Each m_Document As S180Document In m_Product.DocumentList
                    document = doc.CreateElement("Document")
                    documentDescription = doc.CreateElement("Description")
                    documentDescription.InnerText = m_Document.Description
                    documentDate = doc.CreateElement("DocDate")
                    documentDate.InnerText = m_Document.DocDate
                    documentVersion = doc.CreateElement("DocVersion")
                    documentVersion.InnerText = m_Document.DocVersion
                    documentStatus = doc.CreateElement("DocStatus")
                    documentStatus.InnerText = m_Document.DocStatus
                    documentName = doc.CreateElement("DocName")
                    documentName.InnerText = m_Document.DocName
                    documentFileName = doc.CreateElement("FileName")
                    documentFileName.InnerText = m_Document.FileName
                    documentFileType = doc.CreateElement("FileType")
                    documentFileType.InnerText = m_Document.FileType
                    documentContent = doc.CreateElement("Content")
                    documentContent.InnerText = m_Document.Content
                    documentEncodeType = doc.CreateElement("EncodeType")
                    documentEncodeType.InnerText = m_Document.EncodeType
                    documentDocType = doc.CreateElement("DocType")
                    documentDocType.InnerText = m_Document.DocType
                    document.AppendChild(documentDescription)
                    document.AppendChild(documentDate)
                    document.AppendChild(documentVersion)
                    document.AppendChild(documentStatus)
                    document.AppendChild(documentName)
                    document.AppendChild(documentFileName)
                    document.AppendChild(documentFileType)
                    document.AppendChild(documentContent)
                    document.AppendChild(documentEncodeType)
                    document.AppendChild(documentDocType)
                    documentList.AppendChild(document)
                Next
                productList.AppendChild(product)
            Next
            transaction.AppendChild(productList)

            doc.Save(strFullFileName)
            oServiceLog.WriteLogEntry("XML file " + strFullFileName + "created")
            Return "OK"
        Catch ex As Exception
            oServiceLog.WriteLogEntry("Error creating xml file: " + ex.Message)
            Return "Error"
        End Try
    End Function
End Class
