'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Letter
'   Provides the info and collection objects to the client for manipulating
'   an AddressInfo object.
'   Copyright (C) 2004 CIBER, Inc.
'  All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         EN      12/13/04    Original class definition.
'   1.1         EN      12/24/04    Add the properties to collection in Set Method.
'   1.2         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.3         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.4         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Letter" type.
'                                       Also added EntityTypeID attribute to expose the typeID.
'   1.5         AB      02/18/05    Added DataAge check to the Retrieve function
'   1.6         MR      03/27/05    Added Functions to get printed and unprinted Letters.
'
' Function          Description
' Retrieve(ID)     Returns the person/Org requested by the int arg ID
' GetAll()    Returns an  letterCollection with all Person objects
' Add(ID)           Adds the Letter identified by arg ID to the 
'                           internal  letterCollection
' Add(Name)         Adds the Person identified by arg NAME to the internal 
'                    letterCollection            
' Add(Person)       Adds the Person passed as the argument to the internal 
'                            letterCollection
' Remove(ID)        Removes the Person identified by arg ID from the internal 
'                            letterCollection
' Remove(NAME)      Removes the Person identified by arg NAME from the 
'                           internal  letterCollection
' LetterTable()     Returns a datatable containing all columns for the Person 

' colIsDirty()       Returns a boolean indicating whether any of the LetterInfo
'                    objects in the letterCollection has been modified since the
'                    last time it was retrieved from/saved to the repository.
' Flush()            Marshalls all modified/added LetterInfo objects in the 
'                        letterCollection to the repository.
' Save()             Marshalls the internal LetterInfo object to the repository.
Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pLetter
#Region "Public Events"
        Public Event LetterErr(ByVal MsgStr As String)
#End Region
#Region "Private Member Variables"
        Private colLetter As Muster.Info.letterCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private oLetterInfo As Muster.Info.LetterInfo
        Private oLetterDB As New Muster.DataAccess.LetterDB
        Private strNewID As String
        Private nID As Int64 = -1
        Private bolShowDeleted As Boolean
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Letter").ID
        Private pEntityDB As New MUSTER.DataAccess.EntityDB
#End Region
#Region "Constructors"
        Public Sub New()
            oLetterInfo = New MUSTER.Info.LetterInfo
            colLetter = New MUSTER.Info.LetterCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oLetterInfo.ID
            End Get
            Set(ByVal value As Integer)
                oLetterInfo.ID = value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return oLetterInfo.Name
            End Get
            Set(ByVal value As String)
                oLetterInfo.Name = value
                colLetter(oLetterInfo.ID) = oLetterInfo

            End Set
        End Property
        Public Property TypeofDocument() As String
            Get
                Return oLetterInfo.TypeofDocument
            End Get
            Set(ByVal value As String)
                oLetterInfo.TypeofDocument = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property DocumentLocation() As String
            Get
                Return oLetterInfo.DocumentLocation
            End Get
            Set(ByVal value As String)
                oLetterInfo.DocumentLocation = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property EntityType() As Integer
            Get
                Return oLetterInfo.EntityType
            End Get
            Set(ByVal value As Integer)
                oLetterInfo.EntityType = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property EntityId() As Integer
            Get
                Return oLetterInfo.EntityId
            End Get
            Set(ByVal value As Integer)
                oLetterInfo.EntityId = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property DocumentDescription() As String
            Get
                Return oLetterInfo.DocumentDescription
            End Get
            Set(ByVal value As String)
                oLetterInfo.DocumentDescription = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property DatePrinted() As Date
            Get
                Return oLetterInfo.DatePrinted
            End Get
            Set(ByVal value As Date)
                oLetterInfo.DatePrinted = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property WorkFlow() As Integer
            Get
                Return oLetterInfo.WorkFlow
            End Get
            Set(ByVal value As Integer)
                oLetterInfo.WorkFlow = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        'Public ReadOnly Property EntityTypeID() As Integer
        '    Get
        '        Return nEntityTypeID
        '    End Get
        'End Property
        Public Property Deleted() As Boolean
            Get
                Return oLetterInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oLetterInfo.Deleted = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oLetterInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oLetterInfo.IsDirty = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xLetterInfo As MUSTER.Info.LetterInfo
                For Each xLetterInfo In colLetter.Values
                    If xLetterInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Property OwningUser() As String
            Get
                Return oLetterInfo.OwningUser
            End Get
            Set(ByVal value As String)
                oLetterInfo.OwningUser = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property ModuleID() As Integer
            Get
                Return oLetterInfo.ModuleID
            End Get
            Set(ByVal value As Integer)
                oLetterInfo.ModuleID = value
                colLetter(oLetterInfo.ID) = oLetterInfo
            End Set
        End Property
        Public Property EventID() As Int64
            Get
                Return oLetterInfo.EventID
            End Get
            Set(ByVal value As Int64)
                oLetterInfo.EventID = value
            End Set
        End Property
        Public Property EventSequence() As Integer
            Get
                Return oLetterInfo.EventSequence
            End Get
            Set(ByVal Value As Integer)
                oLetterInfo.EventSequence = Value
            End Set
        End Property
        Public Property EventType() As Integer
            Get
                Return oLetterInfo.EventType
            End Get
            Set(ByVal value As Integer)
                oLetterInfo.EventType = value
            End Set
        End Property

        Public ReadOnly Property CreatedBy() As String
            Get
                Return oLetterInfo.CreatedBy
            End Get
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oLetterInfo.CreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedBy() As String
            Get
                Return oLetterInfo.ModifiedBy
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oLetterInfo.ModifiedOn
            End Get
        End Property
        Public Property ShowDeleted() As Boolean
            Get
                Return bolShowDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolShowDeleted = Value
            End Set
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an Address as called for by ID
        Public Function Retrieve(ByVal ID As String) As MUSTER.Info.LetterInfo
            Dim bolDataAged As Boolean = False

            Try
                For Each oLetterInfo In colLetter.Values
                    If oLetterInfo.ID = ID Then
                        If oLetterInfo.IsDirty = False And oLetterInfo.IsAgedData = True Then
                            bolDataAged = True
                        Else
                            Return oLetterInfo
                        End If
                        Exit For
                    End If
                Next
                'If not in memory add it in the database... 
                If bolDataAged Then
                    colLetter.Remove(oLetterInfo)
                End If
                oLetterInfo = oLetterDB.DBGetByID(CLng(ID))
                If oLetterInfo.ID = 0 Then
                    oLetterInfo.ID = nID
                    nID -= 1
                End If
                colLetter.Add(oLetterInfo)
                Return oLetterInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function RetrieveByDocName(ByVal strDocName As String, ByVal strOwningUser As String, Optional ByVal bolDeleted As Boolean = False) As MUSTER.Info.LetterInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oLetterInfo In colLetter.Values
                    If oLetterInfo.Name = strDocName Then
                        If oLetterInfo.IsDirty = False And oLetterInfo.IsAgedData = True Then
                            bolDataAged = True
                        Else
                            Return oLetterInfo
                        End If
                        Exit For
                    End If
                Next
                'If not in memory add it in the database... 
                If bolDataAged Then
                    colLetter.Remove(oLetterInfo)
                End If
                oLetterInfo = oLetterDB.DBGetByDocName(strDocName, strOwningUser, bolDeleted)
                If oLetterInfo.ID = 0 Then
                    oLetterInfo.ID = nID
                    nID -= 1
                End If
                colLetter.Add(oLetterInfo)
                Return oLetterInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Sub SaveDocDescription(ByVal ndocID As Integer, ByVal strdocDesc As String, ByVal isManualDoc As Boolean, ByVal staffID As Integer)
            Try
                oLetterDB.DBSaveDocDescription(ndocID, strdocDesc, ismanualdoc, staffid)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        Function GetAll(Optional ByVal strUserID As String = "", Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.LetterCollection
            Try
                colLetter.Clear()
                colLetter = oLetterDB.GetAllInfo(strUserID, showDeleted)
                Return colLetter
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an letter to the collection as called for by ID
        Public Sub Add(ByVal ID As String)
            Try
                oLetterInfo = oLetterDB.DBGetByID(ID)
                If oLetterInfo.ID = 0 Then
                    oLetterInfo.ID = nID
                    nID -= 1
                End If
                colLetter.Add(oLetterInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Adds an address to the collection as supplied by the caller
        Public Sub Add(ByRef oletter As MUSTER.Info.LetterInfo)
            Try
                oLetterInfo = oletter
                If oLetterInfo.ID = 0 Then
                    oLetterInfo.ID = nID
                    nID -= 1
                End If
                colLetter.Add(oLetterInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the address called for by ID from the collection
        Public Sub Remove(ByVal ID As String)
            Dim myIndex As Int16 = 1
            Try
                For Each oLetterInfo In colLetter.Values
                    If oLetterInfo.ID = ID Then
                        colLetter.Remove(oLetterInfo)
                        'oLetterInfo = New Muster.Info.LetterInfo
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Letter " & ID.ToString & " is not in the collection of Letter.")
        End Sub
        '''Public Function colIsDirty() As Boolean
        '''    Dim oTempInfo As Muster.Info.LetterInfo
        '''    For Each oTempInfo In colLetter.Values
        '''        If oTempInfo.IsDirty Then
        '''            Return True
        '''        End If
        '''    Next
        '''    Return False
        '''End Function
        Public Sub Flush()
            Dim oTempInfo As MUSTER.Info.LetterInfo
            For Each oTempInfo In colLetter.Values
                If oTempInfo.IsDirty Then
                    oLetterInfo = oTempInfo
                    Me.Save()
                End If
            Next
        End Sub
        Public Sub Save()
            Try
                Dim OldKey As String = oLetterInfo.ID.ToString
                oLetterDB.Put(oLetterInfo)
                If oLetterInfo.ID.ToString <> OldKey Then
                    colLetter.ChangeKey(OldKey, oLetterInfo.ID.ToString)
                End If
                oLetterInfo.Archive()
                oLetterInfo.IsDirty = False
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Save(ByRef oLetter As MUSTER.Info.LetterInfo, Optional ByVal Type As String = "SYSTEM")
            Try
                Dim OldKey As String = oLetterInfo.ID.ToString
                oLetterDB.Put(oLetter, Type)
                If oLetterInfo.ID.ToString <> OldKey Then
                    colLetter.ChangeKey(OldKey, oLetterInfo.ID.ToString)
                End If
                oLetterInfo.Archive()
                oLetter.IsDirty = False
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colLetter.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colLetter.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colLetter.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oLetterInfo = New MUSTER.Info.LetterInfo
        End Sub
        Public Sub Reset()
            Dim xLetterInfo As MUSTER.Info.LetterInfo
            If Not colLetter.Values Is Nothing Then
                For Each xLetterInfo In colLetter
                    If xLetterInfo.IsDirty Then
                        xLetterInfo.Reset()
                    End If
                Next
            Else
                xLetterInfo.Reset()
            End If
        End Sub

#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the addresses in the collection
        Public Function PrintedLetterTable(Optional ByVal strUser As String = "") As DataTable
            Dim oletterInfoLocal As MUSTER.Info.LetterInfo
            Dim dr As DataRow
            Dim tbLetterTable As New DataTable
            Dim sList As New SortedList
            Dim i As Integer = 0

            Try

                tbLetterTable.Columns.Add("Type")
                tbLetterTable.Columns.Add("ID")
                tbLetterTable.Columns.Add("Document Type")
                tbLetterTable.Columns.Add("Name")
                tbLetterTable.Columns.Add("Description")
                tbLetterTable.Columns.Add("Date Printed")
                tbLetterTable.Columns.Add("Date Created")
                tbLetterTable.Columns.Add("Document Location")
                tbLetterTable.Columns.Add("WorkFlow")
                tbLetterTable.Columns.Add("EntityID")
                tbLetterTable.Columns.Add("DELETED")
                tbLetterTable.Columns.Add("Created By")
                tbLetterTable.Columns.Add("last_edited_by")
                tbLetterTable.Columns.Add("date_last_edited")
                tbLetterTable.Columns.Add("Entity Type")


                For Each oletterInfoLocal In colLetter.Values
                    If Not oletterInfoLocal.DatePrinted = CDate("01/01/0001") And Not oletterInfoLocal.Deleted And oletterInfoLocal.OwningUser = strUser Then
                        dr = tbLetterTable.NewRow()
                        dr("Type") = pEntityDB.GetEntityTypeDescByID(oletterInfoLocal.EntityType)
                        dr("ID") = oletterInfoLocal.ID
                        dr("Name") = oletterInfoLocal.Name
                        dr("Document Type") = oletterInfoLocal.TypeofDocument
                        dr("Document Location") = oletterInfoLocal.DocumentLocation
                        dr("Entity Type") = oletterInfoLocal.EntityType
                        dr("EntityID") = oletterInfoLocal.EntityId
                        dr("Description") = oletterInfoLocal.DocumentDescription
                        dr("WorkFlow") = oletterInfoLocal.WorkFlow
                        dr("Date Printed") = oletterInfoLocal.DatePrinted.ToShortDateString
                        dr("DELETED") = oletterInfoLocal.Deleted
                        dr("CREATED BY") = oletterInfoLocal.CreatedBy
                        dr("Date Created") = oletterInfoLocal.CreatedOn.ToShortDateString
                        dr("last_edited_by") = oletterInfoLocal.ModifiedBy
                        dr("date_last_edited") = oletterInfoLocal.ModifiedOn.ToShortDateString
                        sList.Add(oletterInfoLocal.ID, dr)
                    End If
                Next
                For i = sList.Count - 1 To 0 Step -1
                    tbLetterTable.Rows.Add(CType(sList.GetByIndex(i), DataRow))
                Next
                Return tbLetterTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function UnPrintedLetterTable(Optional ByVal strUser As String = "") As DataTable
            Dim oletterInfoLocal As MUSTER.Info.LetterInfo
            Dim dr As DataRow
            Dim tbLetterTable As New DataTable
            Dim sList As New SortedList
            Dim i As Integer = 0
            Try

                tbLetterTable.Columns.Add("Type")
                tbLetterTable.Columns.Add("ID")
                tbLetterTable.Columns.Add("Document Type")
                tbLetterTable.Columns.Add("Name")
                tbLetterTable.Columns.Add("Description")
                tbLetterTable.Columns.Add("Date Printed")
                tbLetterTable.Columns.Add("Date Created")
                tbLetterTable.Columns.Add("Document Location")
                tbLetterTable.Columns.Add("WorkFlow")
                tbLetterTable.Columns.Add("EntityID")
                tbLetterTable.Columns.Add("DELETED")
                tbLetterTable.Columns.Add("Created By")
                tbLetterTable.Columns.Add("last_edited_by")
                tbLetterTable.Columns.Add("date_last_edited")
                tbLetterTable.Columns.Add("Entity Type")


                For Each oletterInfoLocal In colLetter.Values
                    If oletterInfoLocal.DatePrinted = CDate("01/01/0001") And Not oletterInfoLocal.Deleted And oletterInfoLocal.OwningUser = strUser Then
                        dr = tbLetterTable.NewRow()
                        dr("Type") = pEntityDB.GetEntityTypeDescByID(oletterInfoLocal.EntityType)
                        dr("ID") = oletterInfoLocal.ID
                        dr("Name") = oletterInfoLocal.Name
                        dr("Document Type") = oletterInfoLocal.TypeofDocument
                        dr("Document Location") = oletterInfoLocal.DocumentLocation
                        dr("Entity Type") = oletterInfoLocal.EntityType
                        dr("EntityID") = oletterInfoLocal.EntityId
                        dr("Description") = oletterInfoLocal.DocumentDescription
                        dr("WorkFlow") = oletterInfoLocal.WorkFlow
                        dr("Date Printed") = oletterInfoLocal.DatePrinted.ToShortDateString
                        dr("DELETED") = oletterInfoLocal.Deleted
                        dr("Created By") = oletterInfoLocal.CreatedBy
                        dr("Date Created") = oletterInfoLocal.CreatedOn.ToShortDateString
                        dr("last_edited_by") = oletterInfoLocal.ModifiedBy
                        dr("date_last_edited") = oletterInfoLocal.ModifiedOn.ToShortDateString
                        sList.Add(oletterInfoLocal.ID, dr)
                    End If
                Next

                For i = sList.Count - 1 To 0 Step -1
                    tbLetterTable.Rows.Add(CType(sList.GetByIndex(i), DataRow))
                Next
                Return tbLetterTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetDataSet(ByVal strSQL As String) As DataSet
            Try
                Dim ds As DataSet
                ds = oLetterDB.DBGetDS(strSQL)
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetDocumentsList(ByVal strUserID As String, Optional ByVal PrintedFlag As Boolean = False) As DataSet
            Dim dsData As DataSet
            Try
                dsData = oLetterDB.GetDocumentsList(strUserID, PrintedFlag)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function UpdatePrintedStatus(ByVal DocumentID As Integer, ByVal DocumentLocation As String, ByVal DatePrinted As DateTime)

            Try
                oLetterDB.UpdatePrintedStatus(DocumentID, DocumentLocation, DatePrinted)
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetCalendarYear(ByVal strUserID As String, Optional ByVal PrintedFlag As Integer = 0)
            Dim dsData As DataSet
            Try
                dsData = oLetterDB.GetCalendarYear(strUserID, PrintedFlag)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function DeleteManualDocuments(ByVal UserID As String)
            Try
                oLetterDB.DeleteManualDocuments(UserID)
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetManualAndSystemDocuments(ByVal strUserID As String) As DataSet
            Dim dsData As DataSet
            Try
                dsData = oLetterDB.GetManualAndSystemDocuments(strUserID)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetManualDocsWithDesc(ByVal strUserID As String) As DataSet
            Dim dsData As DataSet
            Try
                dsData = oLetterDB.GetManualDocsWithDesc(strUserID)
                Return dsData
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
    End Class
End Namespace
