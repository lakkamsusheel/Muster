' -------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pFinancialActivity
' Provides the operations required to manipulate a FinancialActivity object.
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0         AB       06/23/2005    Original class definition
' 
' Function          Description
' -------------------------------------------------------------------------------
' Attribute          Description
' -------------------------------------------------------------------------------
Namespace MUSTER.BusinessLogic
    <Serializable()> Public Class pFinancialActivity

#Region "Private Member Variables"
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        Private colFinActivity As MUSTER.Info.FinancialActivityCollection = New MUSTER.Info.FinancialActivityCollection
        Private oFinActivityInfo As MUSTER.Info.FinancialActivityInfo = New MUSTER.Info.FinancialActivityInfo
        Private oFinActivityDB As MUSTER.DataAccess.FinancialActivityDB = New MUSTER.DataAccess.FinancialActivityDB
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Financial").ID
        Private nID As Int64 = -1
#End Region
#Region "Exposed Events"
        Public Delegate Sub FinActivityBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinActivityBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinActivityBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FinActivityInfoChanged()
        Public Event FinActivityBLChanged As FinActivityBLChangedEventHandler
        Public Event FinActivityBLColChanged As FinActivityBLColChangedEventHandler
        Public Event FinActivityBLErr As FinActivityBLErrEventHandler
        Public Event FinActivityInfChanged As FinActivityInfoChanged
        ' indicates change in the underlying _ProtoInfo structure
#End Region
#Region "Constructors"
        Public Sub New()
            oFinActivityInfo = New MUSTER.Info.FinancialActivityInfo
        End Sub
        Public Sub New(ByVal TextID As Integer)
            oFinActivityInfo = New MUSTER.Info.FinancialActivityInfo
            Me.Retrieve(TextID)
        End Sub

#End Region
#Region "Exposed Attributes"
        Public Property Activity_ID() As Integer
            Get
                Return oFinActivityInfo.ActivityID
            End Get
            Set(ByVal Value As Integer)
                oFinActivityInfo.ActivityID = Value
            End Set
        End Property

        Public Property ActivityDesc() As String
            Get
                Return oFinActivityInfo.ActivityDesc
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.ActivityDesc = Value
            End Set
        End Property

        Public Property ActivityDescShort() As String
            Get
                Return oFinActivityInfo.ActivityDescShort
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.ActivityDescShort = Value
            End Set
        End Property

        Public ReadOnly Property ReimbursementConditionDesc() As String
            Get
                Return oFinActivityInfo.ReimbursementConditionDesc
            End Get
        End Property

        Public Property CostPlus() As Int64
            Get
                Return oFinActivityInfo.CostPlus
            End Get
            Set(ByVal Value As Int64)
                oFinActivityInfo.CostPlus = Value
            End Set
        End Property

        Public Property CostPlusDesc() As String
            Get
                Return oFinActivityInfo.CostPlusDesc
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.CostPlusDesc = Value
            End Set
        End Property

        Public Property FixedPrice() As Int64
            Get
                Return oFinActivityInfo.FixedPrice
            End Get
            Set(ByVal Value As Int64)
                oFinActivityInfo.FixedPrice = Value
            End Set
        End Property


        Public Property FixedPriceDesc() As String
            Get
                Return oFinActivityInfo.FixedPriceDesc
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.FixedPriceDesc = Value
            End Set
        End Property

        Public Property TimeAndMaterials() As Int64
            Get
                Return oFinActivityInfo.TimeAndMaterials
            End Get
            Set(ByVal Value As Int64)
                oFinActivityInfo.TimeAndMaterials = Value
            End Set
        End Property

        Public Property TimeAndMaterialsDesc() As String
            Get
                Return oFinActivityInfo.TimeAndMaterialsDesc
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.TimeAndMaterialsDesc = Value
            End Set
        End Property

        Public Property ReimbursementCondition() As Int64
            Get
                Return oFinActivityInfo.ReimbursementCondition
            End Get
            Set(ByVal Value As Int64)
                oFinActivityInfo.ReimbursementCondition = Value
            End Set
        End Property

        Public Property DueDateStatement() As String
            Get
                Return oFinActivityInfo.DueDateStatement
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.DueDateStatement = Value
            End Set
        End Property

        ' Gets/Sets the active flag for the financial text info object (from info.Active)
        Public Property Active() As Boolean
            Get
                Return oFinActivityInfo.Active
            End Get
            Set(ByVal Value As Boolean)
                oFinActivityInfo.Active = Value
            End Set
        End Property

        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oFinActivityInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.CreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFinActivityInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oFinActivityInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFinActivityInfo.Deleted = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oFinActivityInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFinActivityInfo.IsDirty = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFinActivityInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFinActivityInfo.ModifiedOn
            End Get
        End Property

        Public Property CoverTemplate() As String
            Get
                Return oFinActivityInfo.CoverTemplateDoc
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.CoverTemplateDoc = Value
            End Set
        End Property

        Public Property NoticeTemplate() As String
            Get
                Return oFinActivityInfo.NoticeTemplateDoc
            End Get
            Set(ByVal Value As String)
                oFinActivityInfo.NoticeTemplateDoc = Value
            End Set
        End Property

#End Region
#Region "Exposed Operations"
#Region "General Operations"
        Public Sub Clear()
            oFinActivityInfo = New MUSTER.Info.FinancialActivityInfo
        End Sub
        Public Sub Reset()
            oFinActivityInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        Public Function Retrieve(ByVal TextID As Int64) As MUSTER.Info.FinancialActivityInfo
            Try
                oFinActivityInfo = oFinActivityDB.DBGetByID(TextID)
                If oFinActivityInfo.ActivityID = 0 Then
                    nID -= 1
                End If

                Return oFinActivityInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If Me.ValidateData(strModuleName) Then

                    oFinActivityDB.Put(oFinActivityInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFinActivityInfo.Archive()

                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Financial") As Boolean
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
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xFinancialActivityInfo As MUSTER.Info.FinancialActivityInfo
            Try
                For Each xFinancialActivityInfo In colFinActivity.Values
                    If xFinancialActivityInfo.IsDirty Then
                        oFinActivityInfo = xFinancialActivityInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oFinActivityInfoInfo As MUSTER.Info.FinancialActivityInfo)
            Try
                colFinActivity.Add(oFinActivityInfoInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFinActivityInfo As Object)
            Try
                colFinActivity.Remove(oFinActivityInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        ' Gets all the info
        Public Function GetAll(ByVal nReason As Int64) As MUSTER.Info.FinancialActivityCollection
            Try
                colFinActivity.Clear()
                colFinActivity = oFinActivityDB.DBGetAll
                Return colFinActivity
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        'Inserts new relationship into Financial ACtivity and Tech Doc table
        Public Sub PutFinActivityTechDocRelationship(ByVal ActivityID As Integer, ByVal TecDocID As Integer, ByVal IsSentToFinanceDoc As Boolean, _
                                                     ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                oFinActivityDB.DBPutFinActivityTechDocRelationship(ActivityID, TecDocID, IsSentToFinanceDoc, _
                                                                    moduleID, staffID, returnVal)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub

        'Inserts new relationship into Financial ACtivity and Tech Doc table
        Public Sub ClearACtivityDocRelationshipByID(ByVal ActivityID As Integer, _
                                                     ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                oFinActivityDB.DBRemoveDocsFromActivityID(ActivityID, moduleID, staffID, returnVal)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub

#End Region
#Region " Populate Routines "

        Public Function IsUsed() As Boolean
            Dim dsReturn As New DataSet
            Dim nReturn As Boolean
            Dim strSQL As String
            Try
                strSQL = "select Count(*) as UseCount from tblFIN_COMMITMENT "
                strSQL &= " where  ActivityType = " & oFinActivityInfo.ActivityID

                dsReturn = oFinActivityDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    If dsReturn.Tables(0).Rows(0)("UseCount") > 0 Then
                        nReturn = True
                    Else
                        nReturn = False
                    End If
                Else
                    nReturn = False
                End If
                Return nReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateFinancialActivityList() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialActivityList", False, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateAllTechnicalDocsForActivity() As DataTable
            Try
                Dim dtReturn As DataTable = Me.oFinActivityDB.DBGetAllTechDocsForActivity()
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function PopulateFinancialActivityForCommitment() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialActivityForCommitment", False, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateFinancialAdditionalConditions() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialAdditionalConditions", False, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateReimbursementConditions() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialReimbursementConditions", False, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateCostFormats() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialCostFormats", True, True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateDocumentsForActivity() As DataTable
            Try
                Dim dtReturn As DataTable = Me.oFinActivityDB.DBGetActivityDocs(Activity_ID)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Function GetCostFormatDesign(ByVal costFormat As String) As DataTable
            Try
                Dim dtReturn As DataTable = Me.oFinActivityDB.DBGetCostFormatSpecs(costFormat)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try


        End Function


        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal PropertyMaster As Boolean = True, Optional ByVal IncludeBlank As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            strSQL = ""
            If PropertyMaster Then
                If IncludeBlank Then
                    strSQL = " SELECT '' as PROPERTY_NAME, 0 as PROPERTY_ID, 0 as PROPERTY_POSITION "
                    strSQL &= " UNION "
                End If

                strSQL &= " SELECT PROPERTY_NAME, PROPERTY_ID, PROPERTY_POSITION FROM " & strProperty
                strSQL &= " order by 1 "
            Else
                strSQL &= " SELECT * FROM " & strProperty
            End If
            Try
                dsReturn = oFinActivityDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

#End Region

#End Region
#Region "Private Operations"
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colFinActivity.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.Activity_ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colFinActivity.Item(nArr.GetValue(colIndex + direction)).ActivityID.ToString
            Else
                Return colFinActivity.Item(nArr.GetValue(colIndex)).ActivityID.ToString
            End If
        End Function
#End Region

    End Class
End Namespace


