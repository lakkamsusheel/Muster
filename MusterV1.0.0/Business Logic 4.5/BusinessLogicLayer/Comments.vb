'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Comments
'   Provides the operations required to manipulate an Comments object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         PN      12/14/04    Original class definition.
'   1.1         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.2         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.3         AN      02/02/05    Added Event Handlers for changed event.
'   1.4         JVC2    02/10/2005  Added call to changekey as part of save.
'   1.5         AB      02/17/05    Added DataAge check to the Retrieve function
'                                       Also modified Remove() to remove object outside the For..Loop
'   1.6         MNR     03/22/05    Added Sub Load
'
' Function          Description
' GetComments(Module,EntityType,EntityID)   Returns the Comments requested by the string arg Module,EntityType,EntityID
' GetComments(ID)     Returns the comments requested by the int arg ID
' Add(ID)           Adds the Entity identified by arg ID to the 
'                           internal CommentsCollection
' Add(Comment)         Adds the User identified by arg Comment to the internal 
'                           CommentsCollection
' Add(CommentInfo)       Adds the CommentInfo passed as the argument to the internal 
'                          CommentsCollection
' Remove(ID)       Sets the Comment's deleted property identified by arg ID from the internal 
'                          CommentCollection to true
' Remove(Comment)      Removes the Comment identified by arg Comment from the 
'                           internal CommentCollection
' UserTable()     Returns a datatable containing all columns for the User 
'                           objects in the internal UserCollection.
' UserCombo()     Returns a two-column datatable containing Name and ID for 
'                           the User objects in the internal UserCollection.
'Reset()            Resets the user collection

'-------------------------------------------------------------------------------
Namespace MUSTER.BusinessLogic
    <Serializable()> _
Public Class pComments
#Region "Public Events"
        'Public Event CommentExists(ByVal MsgStr As String)
        Public Event InfoBecameDirty(ByVal BolValue As Boolean)
        'added by kiran
        Public Event evtCommentColOwner(ByVal Ownerid As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        Public Event evtCommentColFac(ByVal Facid As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        Public Event evtCommentColTanks(ByVal Tankid As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        Public Event evtCommentColPipe(ByVal pipeid As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        Public Event evtCommentColLustActivity(ByVal ActivityID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        Public Event evtCommentColLustDocument(ByVal DocumentID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)
        Public Event evtCommentColClosureEvent(ByVal closureID As Integer, ByVal commentsCol As MUSTER.Info.CommentsCollection)

        'end changes
        Public Event evtCommentInfoOwner(ByVal OwnerID As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo)
        Public Event evtCommentInfoFac(ByVal Facid As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo)
        Public Event evtCommentInfoTanks(ByVal Tankid As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo)
        Public Event evtCommentInfoPipe(ByVal pipeid As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo)
        Public Event evtCommentInfoLustActivity(ByVal ActivityID As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo)
        Public Event evtCommentInfoLustDocument(ByVal DocumentID As Integer, ByVal commentsInfo As MUSTER.Info.CommentsInfo)
#End Region

#Region "Private member variables"
        Private colComments As MUSTER.Info.CommentsCollection
        Private WithEvents oCommentsInfo As MUSTER.Info.CommentsInfo
        Private oCommentsDB As New MUSTER.DataAccess.CommentsDB
        Private nID As Integer = -1
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            colComments = New MUSTER.Info.CommentsCollection
            oCommentsInfo = New MUSTER.Info.CommentsInfo
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oCommentsInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oCommentsInfo.ID = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oCommentsInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oCommentsInfo.Deleted = Value
            End Set
        End Property
        Public Property EntityID() As Integer
            Get
                Return oCommentsInfo.EntityID
            End Get
            Set(ByVal Value As Integer)
                oCommentsInfo.EntityID = Value
            End Set
        End Property
        Public Property EntityAdditionalInfo() As String
            Get
                Return oCommentsInfo.EntityAdditionalInfo
            End Get
            Set(ByVal Value As String)
                oCommentsInfo.EntityAdditionalInfo = Value
            End Set
        End Property
        Public Property EntityType() As Integer
            Get
                Return oCommentsInfo.EntityType
            End Get
            Set(ByVal Value As Integer)
                oCommentsInfo.EntityType = Value
            End Set
        End Property

        Public Property CommentsScope() As String
            Get
                Return oCommentsInfo.CommentsScope
            End Get
            Set(ByVal Value As String)
                oCommentsInfo.CommentsScope = Value
            End Set
        End Property

        Public Property UserID() As String
            Get
                Return oCommentsInfo.UserID
            End Get
            Set(ByVal Value As String)
                oCommentsInfo.UserID = Value
            End Set
        End Property
        Public Property ModuleName() As String
            Get
                Return oCommentsInfo.ModuleName
            End Get
            Set(ByVal Value As String)
                oCommentsInfo.ModuleName = Value
            End Set
        End Property
        Public Property WorkFlow() As String
            Get
                Return oCommentsInfo.WorkFlow
            End Get
            Set(ByVal Value As String)
                oCommentsInfo.WorkFlow = Value
            End Set
        End Property
        Public Property Comments() As String
            Get
                Return oCommentsInfo.Comments
            End Get
            Set(ByVal Value As String)
                oCommentsInfo.Comments = Value
            End Set
        End Property
        Public Property CommentDate() As Date
            Get
                Return oCommentsInfo.CommentDate
            End Get
            Set(ByVal Value As Date)
                oCommentsInfo.CommentDate = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                If oCommentsInfo.IsDirty Or colIsDirty Then
                    Return True
                Else
                    Return False
                End If
            End Get

            Set(ByVal value As Boolean)
                oCommentsInfo.IsDirty = value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oCommentsInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oCommentsInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oCommentsInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oCommentsInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oCommentsInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oCommentsInfo.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByVal ds As DataSet, ByVal ModuleName As String, ByVal EntityType As Integer, ByVal entityID As Integer, Optional ByVal entityAddnInfo As String = "")
            Dim dr As DataRow
            Try
                For Each dr In ds.Tables("Comments").Rows
                    If dr.Item("MODULE_ID") = ModuleName And _
                        dr.Item("ENTITY_TYPE") = EntityType And _
                        dr.Item("ENTITY_ID") = entityID And _
                        dr.Item("ENTITY_ADDITIONAL_INFO") = IIf(entityAddnInfo = String.Empty, dr.Item("ENTITY_ADDITIONAL_INFO"), entityAddnInfo) Then
                        oCommentsInfo = New MUSTER.Info.CommentsInfo(dr)
                        Select Case (EntityType)
                            Case 6
                                RaiseEvent evtCommentInfoFac(entityID, oCommentsInfo)
                            Case 9
                                RaiseEvent evtCommentInfoOwner(entityID, oCommentsInfo)
                            Case 10
                                RaiseEvent evtCommentInfoPipe(entityID, oCommentsInfo)
                            Case 12
                                RaiseEvent evtCommentInfoTanks(entityID, oCommentsInfo)
                            Case 23
                                RaiseEvent evtCommentInfoLustActivity(entityID, oCommentsInfo)
                            Case 24
                                RaiseEvent evtCommentInfoLustDocument(entityID, oCommentsInfo)
                        End Select
                        colComments.Add(oCommentsInfo)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        'Obtains and returns an Comment as called for by ID
        Public Function Retrieve(ByVal nCommentID As Integer, Optional ByVal strUserID As String = "") As MUSTER.Info.CommentsInfo
            Dim oCommentsInfoLocal As MUSTER.Info.CommentsInfo
            Dim bolDataAged As Boolean = False

            Try
                For Each oCommentsInfoLocal In colComments.Values
                    If oCommentsInfoLocal.ID = nCommentID And _
                        oCommentsInfoLocal.UserID = IIf(strUserID = "", oCommentsInfoLocal.UserID, strUserID) Then
                        If oCommentsInfoLocal.IsAgedData = True And oCommentsInfoLocal.IsDirty = False Then
                            bolDataAged = True
                        Else
                            oCommentsInfo = oCommentsInfoLocal
                            Return oCommentsInfo
                        End If
                    End If
                Next
                If bolDataAged Then
                    colComments.Remove(oCommentsInfoLocal)
                End If
                oCommentsInfo = oCommentsDB.DBGetByID(nCommentID, strUserID)
                colComments.Add(oCommentsInfo)

                Return oCommentsInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                Dim nCommentID As Int64
                Dim oldID As Int64 = Me.ID
                nCommentID = oCommentsDB.put(oCommentsInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Function
                End If
                If oldID <> Me.ID Then
                    colComments.ChangeKey(oldID, Me.ID)
                End If
                'If oCommentsInfo.ID <> nCommentID Then
                '    oCommentsInfo.ID = nCommentID
                'End If

                oCommentsInfo.IsDirty = False
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.CommentsCollection
            Try
                colComments.Clear()
                colComments = oCommentsDB.DBGetAllCommentsInfo
                Return colComments
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Function GetByModule(ByVal ModuleName As String, ByVal EntityType As Integer, ByVal entityID As Integer, Optional ByVal entityAddnInfo As String = "") As DataTable
        '    Try
        '        colComments = oCommentsDB.DBGetByModuleName(ModuleName, EntityType, entityID, entityAddnInfo)
        '        'added by kiran
        '        If colComments.Count > 0 Then
        '            ' raiseevent to facility (entityID- facility ID) with col
        '            If (EntityType = 9) Then
        '                RaiseEvent evtCommentColOwner(entityID, colComments)
        '            ElseIf (EntityType = 6) Then
        '                RaiseEvent evtCommentColFac(entityID, colComments)
        '            ElseIf (EntityType = 12) Then
        '                RaiseEvent evtCommentColTanks(entityID, colComments)
        '            ElseIf (EntityType = 10) Then
        '                RaiseEvent evtCommentColPipe(entityID, colComments)
        '            ElseIf (EntityType = 23) Then
        '                RaiseEvent evtCommentColLustActivity(entityID, colComments)
        '            ElseIf (EntityType = 24) Then
        '                RaiseEvent evtCommentColLustDocument(entityID, colComments)
        '            ElseIf (EntityType = 22) Then
        '                RaiseEvent evtCommentColClosureEvent(entityID, colComments)
        '            End If
        '        End If
        '        'end changes
        '        Return CommentsTable()
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        'Function GetByEntityType(ByVal EntityType As Integer, ByVal entityID As Integer, Optional ByVal entityAddnInfo As String = "") As DataTable
        '    Try
        '        colComments = oCommentsDB.DBGetByEntityType(EntityType, entityID, entityAddnInfo)
        '        'added by kiran
        '        If colComments.Count > 0 Then
        '            ' raiseevent to facility (entityID- facility ID) with col
        '            If (EntityType = 9) Then
        '                RaiseEvent evtCommentColOwner(entityID, colComments)
        '            ElseIf (EntityType = 6) Then
        '                RaiseEvent evtCommentColFac(entityID, colComments)
        '            ElseIf (EntityType = 12) Then
        '                RaiseEvent evtCommentColTanks(entityID, colComments)
        '            ElseIf (EntityType = 10) Then
        '                RaiseEvent evtCommentColPipe(entityID, colComments)
        '            ElseIf (EntityType = 23) Then
        '                RaiseEvent evtCommentColLustActivity(entityID, colComments)
        '            ElseIf (EntityType = 24) Then
        '                RaiseEvent evtCommentColLustDocument(entityID, colComments)
        '            ElseIf (EntityType = 22) Then
        '                RaiseEvent evtCommentColClosureEvent(entityID, colComments)
        '            End If
        '        End If
        '        'end changes
        '        Return CommentsTable()
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        Function GetComments(Optional ByVal strModuleName As String = "", Optional ByVal entityType As Integer = 0, Optional ByVal entityID As Integer = 0, Optional ByVal entityAddnInfo As String = "", Optional ByVal userID As String = "", Optional ByVal commentID As Integer = 0, Optional ByVal showDeleted As Boolean = False) As DataSet
            Return oCommentsDB.DBGetComments(strModuleName, entityType, entityID, entityAddnInfo, userID, commentID, showDeleted)
        End Function
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oComments As MUSTER.Info.CommentsInfo)
            Try
                oCommentsInfo = oComments
                If oCommentsInfo.ID = 0 Then
                    oComments.ID = nID
                    nID -= 1
                End If
                colComments.Add(oCommentsInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Sets the deleted property to true in the collection
        Public Sub Remove(ByVal ID As Int64)

            Dim bolFoundRecord As Boolean = False
            Dim xCommentsInfo As MUSTER.Info.CommentsInfo

            Try
                xCommentsInfo = colComments.Item(ID)
                'For Each xCommentsInfo In colComments.Values
                '    If xCommentsInfo.ID = ID Then
                '        'oCommentsInfo = xCommentsInfo
                '        'oCommentsInfo.Deleted = True
                '        'TODO Check with Jay
                '        'Added Collection Remove
                '        'colComments.Remove(xCommentsInfo)

                '        bolFoundRecord = True
                '        Exit For
                '    End If
                'Next

                If Not (xCommentsInfo Is Nothing) Then
                    colComments.Remove(xCommentsInfo)
                End If
                'If bolFoundRecord Then
                '    colComments.Remove(xCommentsInfo)
                'End If

                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Comment " & ID.ToString & " is not in the collection of Comments.")

        End Sub
        'Removes the Comment supplied from the collection
        Public Sub Remove(Optional ByVal oCommentsInf As MUSTER.Info.CommentsInfo = Nothing)

            Try
                If Not oCommentsInf Is Nothing Then
                    oCommentsInfo = oCommentsInf
                End If
                colComments.Remove(oCommentsInfo)
                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Comments " & oCommentsInf.Comments & " is not in the collection of Comments.")

        End Sub
        Private Property colIsDirty() As Boolean
            Get
                Dim xCommentsInfo As MUSTER.Info.CommentsInfo
                For Each xCommentsInfo In colComments.Values
                    If xCommentsInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xCommentsInfo As MUSTER.Info.CommentsInfo
            For Each xCommentsInfo In colComments.Values
                If xCommentsInfo.IsDirty Then
                    oCommentsInfo = xCommentsInfo
                    Me.Save(moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                End If
            Next
            'Adam Nall - Added to change the pProfile isDirty. It changes this on the save 
            '            of the single profileInfo but not the parent class isDirty
            Me.IsDirty = False
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colComments.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colComments.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colComments.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Reset()
            Dim oCommentsInfoLocal As MUSTER.Info.CommentsInfo
            For Each oCommentsInfoLocal In colComments.Values
                oCommentsInfoLocal.Reset()
                oCommentsInfoLocal.IsDirty = False
            Next
            If colComments Is Nothing Then
                oCommentsInfo.Reset()
            End If
        End Sub

        Public Sub Clear(Optional ByVal strDepth As String = "ALL")
            Me.colComments = New MUSTER.Info.CommentsCollection
            Me.oCommentsInfo = New MUSTER.Info.CommentsInfo
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function CommentsTable() As DataTable
            Dim drRow As DataRow
            Dim dtComments As New DataTable
            Dim xCommentsInfo As MUSTER.Info.CommentsInfo
            Try
                dtComments.Columns.Add("MODULE")
                dtComments.Columns.Add("USER ID")
                dtComments.Columns.Add("CREATED ON", GetType(System.DateTime))
                dtComments.Columns.Add("CREATEDBY")
                dtComments.Columns.Add("COMMENTS")
                dtComments.Columns.Add("VIEWABLE BY")
                dtComments.Columns.Add("ENTITY_ID")
                dtComments.Columns.Add("ENTITY_TYPE")
                dtComments.Columns.Add("DELETED", GetType(System.Boolean))
                dtComments.Columns.Add("COMMENT_ID")
                dtComments.Columns.Add("ENTITY_ADDITIONAL_INFO")
                dtComments.Columns("COMMENT_ID").DefaultValue = -1
                dtComments.Columns("DELETED").DefaultValue = False
                dtComments.Columns("CREATED ON").DefaultValue = System.DateTime.Today.ToShortDateString

                For Each xCommentsInfo In colComments.Values
                    drRow = dtComments.NewRow

                    drRow("MODULE") = xCommentsInfo.ModuleName
                    drRow("USER ID") = xCommentsInfo.UserID
                    drRow("CREATED ON") = xCommentsInfo.CommentDate
                    drRow("CREATEDBY") = xCommentsInfo.UserID
                    drRow("COMMENTS") = xCommentsInfo.Comments
                    drRow("VIEWABLE BY") = xCommentsInfo.CommentsScope
                    drRow("COMMENT_ID") = xCommentsInfo.ID
                    drRow("DELETED") = xCommentsInfo.Deleted
                    drRow("ENTITY_ID") = xCommentsInfo.EntityID
                    drRow("ENTITY_TYPE") = xCommentsInfo.EntityType
                    drRow("ENTITY_ADDITIONAL_INFO") = xCommentsInfo.EntityAdditionalInfo
                    dtComments.Rows.Add(drRow)
                Next
                Return dtComments
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

#End Region
#End Region
#Region "External Event Handlers"
        Private Sub ThisInfo(ByVal bolValue As Boolean) Handles oCommentsInfo.CommentsInfoChanged
            '
            ' Alert the client that the current info object data has changed
            '
            RaiseEvent InfoBecameDirty(Me.IsDirty Or Me.colIsDirty)
        End Sub
#End Region
    End Class
End Namespace
