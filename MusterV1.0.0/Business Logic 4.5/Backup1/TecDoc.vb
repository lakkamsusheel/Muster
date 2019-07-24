
Namespace MUSTER.BusinessLogic
    <Serializable()> Public Class pTecDoc
        '-------------------------------------------------------------------------------
        ' MUSTER.BusinessLogic.TechDoc
        '   Provides the operations required to manipulate a Technical Document object.
        '
        ' Copyright (C) 2004, 2005 CIBER, Inc.
        ' All rights reserved.
        '
        ' Release   Initials    Date        Description
        '  1.0         JC       5/24/2005    Original class definition
        '
        ' Function          Description
        '-------------------------------------------------------------------------------
        ' Attribute          Description
        '-------------------------------------------------------------------------------
#Region "Private Member Variables"
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private colTecDocs As MUSTER.Info.TecDocCollection = New MUSTER.Info.TecDocCollection
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Tech Doc").ID
        Private nID As Int64 = -1
        Private oTecDocDB As MUSTER.DataAccess.TecDocDB = New MUSTER.DataAccess.TecDocDB
        Private oTecDocInfo As MUSTER.Info.TecDocInfo = New MUSTER.Info.TecDocInfo
#End Region
#Region "Public Events"
        Public Delegate Sub TecDocBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub TecDocChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub TecDocErrEventHandler(ByVal MsgStr As String)
        ' indicates change in the underlying TecDocInfo structure
        Public Delegate Sub TecDocInfoChanged()

        Public Event TecDocChanged As TecDocChangedEventHandler
        Public Event TecDocColChanged As TecDocBLColChangedEventHandler
        Public Event TecDocErr As TecDocErrEventHandler
#End Region
#Region "Constructors"
        Public Sub New()
            oTecDocInfo = New MUSTER.Info.TecDocInfo
        End Sub
        Public Sub New(ByVal DocumentID As Integer)
            oTecDocInfo = New MUSTER.Info.TecDocInfo

            Me.Retrieve(DocumentID)
        End Sub
        Public Sub New(ByVal LustEventName As String)
            oTecDocInfo = New MUSTER.Info.TecDocInfo
            Me.Retrieve(LustEventName)
        End Sub
#End Region
#Region "Exposed Attributes"
        ' Gets/Sets the active flag for the technical document (from TecDoc.Active)
        Public Property Active() As Boolean
            Get
                Return oTecDocInfo.Active()
            End Get
            Set(ByVal Value As Boolean)
                oTecDocInfo.Active = Value
            End Set
        End Property
        ' Gets/Sets the 1st automatically generated technical document associated with this technical document (from oTecDoc.Auto_Doc_1)
        Public Property Auto_Doc_1() As Long
            Get
                Return oTecDocInfo.Auto_Doc_1
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_1 = Value
            End Set
        End Property
        ' Gets/Sets the 2nd automatically generated technical document associated with this technical document (from oTecDoc.Auto_Doc_2)
        Public Property Auto_Doc_2() As Long
            Get
                Return oTecDocInfo.Auto_Doc_2
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_2 = Value
            End Set
        End Property

        Public Property FinActivityType() As Integer
            Get
                Return oTecDocInfo.FinActivityType
            End Get
            Set(ByVal Value As Integer)
                oTecDocInfo.FinActivityType = Value
            End Set
        End Property

        ' Gets/Sets the 3rd automatically generated technical document associated with this technical document (from oTecDoc.Auto_Doc_3)
        Public Property Auto_Doc_3() As Long
            Get
                Return oTecDocInfo.Auto_Doc_3
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_3 = Value
            End Set
        End Property
        ' Gets/Sets the 4th automatically generated technical document associated with this technical document (from oTecDoc.Auto_Doc_4)
        Public Property Auto_Doc_4() As Long
            Get
                Return oTecDocInfo.Auto_Doc_4
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_4 = Value
            End Set
        End Property
        ' The 5th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_5() As Long
            Get
                Return oTecDocInfo.Auto_Doc_5
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_5 = Value
            End Set
        End Property
        ' The 6th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_6() As Long
            Get
                Return oTecDocInfo.Auto_Doc_6
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_6 = Value
            End Set
        End Property
        ' The 7th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_7() As Long
            Get
                Return oTecDocInfo.Auto_Doc_7
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_7 = Value
            End Set
        End Property
        ' The 8th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_8() As Long
            Get
                Return oTecDocInfo.Auto_Doc_8
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_8 = Value
            End Set
        End Property
        ' The 9th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_9() As Long
            Get
                Return oTecDocInfo.Auto_Doc_9
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_9 = Value
            End Set
        End Property
        ' The 10th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_10() As Long
            Get
                Return oTecDocInfo.Auto_Doc_10
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.Auto_Doc_10 = Value
            End Set
        End Property

        Public Property colIsDirty() As Boolean
            Get
                Dim xTecDocInfo As MUSTER.Info.TecDocInfo
                For Each xTecDocInfo In colTecDocs.Values
                    If xTecDocInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oTecDocInfo.IsDirty = Value
            End Set
        End Property
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oTecDocInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oTecDocInfo.CreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oTecDocInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oTecDocInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oTecDocInfo.Deleted = Value
            End Set
        End Property
        ' The entity ID associated with a technical document.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return oTecDocInfo.EntityID
            End Get
        End Property
        ' Gets/Sets the physical file name of the technical document (from oTecDoc.Physical_File_Name)
        Public Property FileName() As String
            Get
                Return oTecDocInfo.Physical_File_Name
            End Get
            Set(ByVal Value As String)
                oTecDocInfo.Physical_File_Name = Value
            End Set
        End Property
        ' Gets the ID of the technical document (from oTecDoc.ID)
        Public Property ID() As Long
            Get
                Return oTecDocInfo.ID
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.ID = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oTecDocInfo.IsDirty()
            End Get
            Set(ByVal Value As Boolean)
                oTecDocInfo.IsDirty = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oTecDocInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oTecDocInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oTecDocInfo.ModifiedOn
            End Get
        End Property
        ' Gets/Sets the name of the technical document (from oTecDoc.Name)
        Public Property Name() As String
            Get
                Return oTecDocInfo.Name
            End Get
            Set(ByVal Value As String)
                oTecDocInfo.Name = Value
            End Set
        End Property
        ' Gets/Sets the NTFE/EUD flag (from oTecDoc.NTFE_Flag)
        Public Property NTFE_Flag() As Boolean
            Get
                Return oTecDocInfo.NTFE_Flag
            End Get
            Set(ByVal Value As Boolean)
                oTecDocInfo.NTFE_Flag = Value
            End Set
        End Property
        ' Gets/Sets the STFS/STFS-Direc/Federal flag (from oTecDoc.STFS_Flag)
        Public Property STFS_Flag() As Boolean
            Get
                Return oTecDocInfo.STFS_Flag
            End Get
            Set(ByVal Value As Boolean)
                oTecDocInfo.STFS_Flag = Value
            End Set
        End Property

        ' Gets/Sets the technical document type (from oTecDoc.DocType)
        Public Property DocType() As Long
            Get
                Return oTecDocInfo.DocType
            End Get
            Set(ByVal Value As Long)
                oTecDocInfo.DocType = Value
            End Set
        End Property
        ' Gets/Sets the trigger field for the technical document (from oTecDoc.Trigger_Field)
        Public Property Trigger_Field() As Int64
            Get
                Return oTecDocInfo.Trigger_Field
            End Get
            Set(ByVal Value As Int64)
                oTecDocInfo.Trigger_Field = Value
            End Set
        End Property
        ' Gets/Sets the name of the technical document (from oTecDoc.Name)
        Public ReadOnly Property InfoObject() As MUSTER.Info.TecDocInfo
            Get
                Return oTecDocInfo
            End Get
        End Property
#End Region
#Region "Exposed Methods"
#Region "General Operations"
        Public Sub Clear()
            oTecDocInfo = New MUSTER.Info.TecDocInfo
        End Sub
        Public Sub Reset()
            oTecDocInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal DocumentID As Int64) As MUSTER.Info.TecDocInfo
            Try
                oTecDocInfo = oTecDocDB.DBGetByID(DocumentID)
                If oTecDocInfo.ID = 0 Then
                    nID -= 1
                End If

                Return oTecDocInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal DocumentName As String) As MUSTER.Info.TecDocInfo

            Try
                oTecDocInfo = oTecDocDB.DBGetByName(DocumentName)
                If oTecDocInfo.ID = 0 Then
                    nID -= 1
                End If

                Return oTecDocInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim strModuleName As String = String.Empty
            Dim bolSubmitForCalendar As Boolean
            Try
                If Me.ValidateData(strModuleName) Then

                    oTecDocDB.Put(oTecDocInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oTecDocInfo.Archive()
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True

            Try
                Select Case [module]
                    Case "Registration"

                        ' if any validations failed
                        Exit Select
                    Case "Technical"

                        ' if any validations failed
                        Exit Select
                End Select
                'If errStr.Length > 0 Or Not validateSuccess Then
                '    RaiseEvent LustEventErr(errStr)
                'End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

#End Region
#Region "Collection Operations"

        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oTecDocInfo As MUSTER.Info.TecDocInfo)
            Try
                colTecDocs.Add(oTecDocInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        ' Gets all the info
        Public Function GetAll() As MUSTER.Info.TecDocCollection
            Try
                colTecDocs.Clear()
                colTecDocs = oTecDocDB.GetAllInfo
                Return colTecDocs
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xTecDocInfo As MUSTER.Info.TecDocInfo
            Try
                For Each xTecDocInfo In colTecDocs.Values
                    If xTecDocInfo.IsDirty Then
                        oTecDocInfo = xTecDocInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Sub

        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oTecDocInfo)
            Try
                colTecDocs.Remove(oTecDocInfo)
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
            Dim strArr() As String = colTecDocs.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colTecDocs.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colTecDocs.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region " Populate Routines "
        Public Function GetCantNFADocID() As Int64
            Dim nReturn As Int64
            Dim strSQL As String
            Try

                strSQL = "Select top 1 Document_ID from tblTEC_DOCUMENT where DocName like 'Can%t%NFA%' and deleted = 0 "

                Dim dtTemp As DataTable = GetDataTable(strSQL, False, True)
                If dtTemp.Rows.Count > 0 Then
                    nReturn = dtTemp.Rows(0)("Document_ID")
                Else
                    nReturn = 0
                End If
                Return nReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetLustEventUsage() As Int64
            Dim nReturn As Int64
            Dim strSQL As String
            Try

                strSQL = "Select Document_ID, "
                strSQL &= "isnull((select count(*) from dbo.tblTEC_ACT_DOC_RELATIONSHIP where Activity_ID in (Select Activity_ID from dbo.tblTEC_ACTIVITY where Active = 1 and Deleted = 0) and Document_ID = tblTEC_DOCUMENT.Document_ID),0) as ActivityUsage,"
                strSQL &= "isnull((select count(*) from dbo.tblTEC_EVENT_ACTIVITY_DOCUMENT where Deleted = 0 and Document_Property_ID = tblTEC_DOCUMENT.Document_ID),0) as LustEventUsage "
                strSQL &= "from dbo.tblTEC_DOCUMENT where Document_ID = " & Me.ID.ToString

                Dim dtTemp As DataTable = GetDataTable(strSQL, False, True)
                If dtTemp Is Nothing Then
                    nReturn = 0
                ElseIf dtTemp.Rows.Count > 0 Then
                    nReturn = dtTemp.Rows(0)("LustEventUsage")
                Else
                    nReturn = 0
                End If
                Return nReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetActivityUsage() As Int64
            Dim nReturn As Int64
            Dim strSQL As String
            Try

                strSQL = "Select Document_ID, "
                strSQL &= "isnull((select count(*) from dbo.tblTEC_ACT_DOC_RELATIONSHIP where Activity_ID in (Select Activity_ID from dbo.tblTEC_ACTIVITY where Active = 1 and Deleted = 0) and Document_ID = tblTEC_DOCUMENT.Document_ID),0) as ActivityUsage,"
                strSQL &= "isnull((select count(*) from dbo.tblTEC_EVENT_ACTIVITY_DOCUMENT where Deleted = 0 and Document_Property_ID = tblTEC_DOCUMENT.Document_ID),0) as LustEventUsage "
                strSQL &= "from dbo.tblTEC_DOCUMENT where Document_ID = " & Me.ID.ToString

                Dim dtTemp As DataTable = GetDataTable(strSQL, False, True)
                If dtTemp Is Nothing Then
                    nReturn = 0
                ElseIf dtTemp.Rows.Count > 0 Then
                    nReturn = dtTemp.Rows(0)("ActivityUsage")
                Else
                    nReturn = 0
                End If
                Return nReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateTecDocumentTriggerList(Optional ByVal IncludeBlank As Boolean = False) As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vTECDocumentTriggers", IncludeBlank, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateTecDocumentList(Optional ByVal IncludeBlank As Boolean = False, Optional ByVal ShowAll As Boolean = True) As DataTable
            Dim dtReturn As DataTable
            Try
                If ShowAll Then
                    dtReturn = GetDataTable("VTECDocumentList", IncludeBlank, False)
                Else
                    dtReturn = GetDataTable("VTECDocumentList_Active", IncludeBlank, False)
                End If

                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateTecDocumentTypes() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VTECDocumentTypes", False, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal IncludeBlank As Boolean = False, Optional ByVal CustomSQL As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            strSQL = ""
            If CustomSQL Then
                strSQL = strProperty
            Else
                If IncludeBlank Then
                    strSQL = " SELECT '' as PROPERTY_NAME, 0 as PROPERTY_ID "
                    strSQL &= " UNION "
                End If

                strSQL &= " SELECT PROPERTY_NAME, PROPERTY_ID FROM " & strProperty
                strSQL &= " order by 1 "
            End If
            Try
                dsReturn = oTecDocDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetAutoCreatedDocumentParent(ByVal nEventActivityID As Integer, ByVal nAutoDocId As Integer) As Integer

            Try
                Return oTecDocDB.DBGetAutoCreatedDocumentParent(nEventActivityID, nAutoDocId)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
    End Class
End Namespace
