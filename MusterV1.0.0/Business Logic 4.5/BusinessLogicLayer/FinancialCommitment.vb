' -------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pFinancialCommitment
' Provides the operations required to manipulate a FinancialCommitment object.
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0         AB       06/24/2005    Original class definition
'
' 2.0   Thomas Franey  02/24/2009    Aded osInstallSetup & nInstallSetup to Info Object
' 
' Function          Description
' -------------------------------------------------------------------------------
' Attribute          Description
' -------------------------------------------------------------------------------
Namespace MUSTER.BusinessLogic

    <Serializable()> Public Class pFinancialCommitment

#Region "Private Member Variables"
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        Private colFinancialCommitment As MUSTER.Info.FinancialCommitmentCollection = New MUSTER.Info.FinancialCommitmentCollection
        Private oFinancialCommitmentInfo As MUSTER.Info.FinancialCommitmentInfo = New MUSTER.Info.FinancialCommitmentInfo
        Private oFinancialCommitmentDB As MUSTER.DataAccess.FinancialCommitmentDB = New MUSTER.DataAccess.FinancialCommitmentDB
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Financial").ID
        Private nID As Int64 = -1
#End Region
#Region "Exposed Events"
        Public Delegate Sub FinancialCommitmentBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialCommitmentBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialCommitmentBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FinancialCommitmentInfoChanged()
        Public Event FinancialCommitmentBLChanged As FinancialCommitmentBLChangedEventHandler
        Public Event FinancialCommitmentBLColChanged As FinancialCommitmentBLColChangedEventHandler
        Public Event FinancialCommitmentBLErr As FinancialCommitmentBLErrEventHandler
        Public Event FinancialCommitmentInfChanged As FinancialCommitmentInfoChanged
        ' indicates change in the underlying _ProtoInfo structure
#End Region
#Region "Constructors"
        Public Sub New()
            oFinancialCommitmentInfo = New MUSTER.Info.FinancialCommitmentInfo
        End Sub
        Public Sub New(ByVal TextID As Integer)
            oFinancialCommitmentInfo = New MUSTER.Info.FinancialCommitmentInfo
            Me.Retrieve(TextID)
        End Sub

#End Region
#Region "Exposed Attributes"
        ' The unique ID for the row containing the text in the table (info.ID)
        Public ReadOnly Property CommitmentID() As Int64
            Get
                Return oFinancialCommitmentInfo.CommitmentID
            End Get
        End Property

        Public Property Fin_Event_ID() As Int64
            Get
                Return oFinancialCommitmentInfo.Fin_Event_ID
            End Get
            Set(ByVal Value As Int64)
                oFinancialCommitmentInfo.Fin_Event_ID = Value
            End Set
        End Property
        Public Property FundingType() As Int64
            Get
                Return oFinancialCommitmentInfo.FundingType
            End Get
            Set(ByVal Value As Int64)
                oFinancialCommitmentInfo.FundingType = Value
            End Set
        End Property

        Public Property Comments() As String
            Get
                Return oFinancialCommitmentInfo.Comments
            End Get
            Set(ByVal Value As String)
                oFinancialCommitmentInfo.Comments = Value
            End Set
        End Property
        Public Property PONumber() As String
            Get
                Return oFinancialCommitmentInfo.PONumber
            End Get
            Set(ByVal Value As String)
                oFinancialCommitmentInfo.PONumber = Value
            End Set
        End Property
        Public Property NewPONumber() As String
            Get
                Return oFinancialCommitmentInfo.NewPONumber
            End Get
            Set(ByVal Value As String)
                oFinancialCommitmentInfo.NewPONumber = Value
            End Set
        End Property
        Public Property RollOver() As Boolean
            Get
                Return oFinancialCommitmentInfo.RollOver
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitmentInfo.RollOver = Value
            End Set
        End Property
        Public Property ZeroOut() As Boolean
            Get
                Return oFinancialCommitmentInfo.ZeroOut
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitmentInfo.ZeroOut = Value
            End Set
        End Property
        Public Property ApprovedDate() As Date
            Get
                Return oFinancialCommitmentInfo.ApprovedDate
            End Get
            Set(ByVal Value As Date)
                oFinancialCommitmentInfo.ApprovedDate = Value
            End Set
        End Property
        Public Property SOWDate() As Date
            Get
                Return oFinancialCommitmentInfo.SOWDate
            End Get
            Set(ByVal Value As Date)
                oFinancialCommitmentInfo.SOWDate = Value
            End Set
        End Property
        Public Property ContractType() As Int64
            Get
                Return oFinancialCommitmentInfo.ContractType
            End Get
            Set(ByVal Value As Int64)
                oFinancialCommitmentInfo.ContractType = Value
            End Set
        End Property
        Public Property ActivityType() As Int64
            Get
                Return oFinancialCommitmentInfo.ActivityType
            End Get
            Set(ByVal Value As Int64)
                oFinancialCommitmentInfo.ActivityType = Value
            End Set
        End Property
        Public Property ReimbursementCondition() As String
            Get
                Return oFinancialCommitmentInfo.ReimbursementCondition
            End Get
            Set(ByVal Value As String)
                oFinancialCommitmentInfo.ReimbursementCondition = Value
            End Set
        End Property
        Public Property DueDate() As Date
            Get
                Return oFinancialCommitmentInfo.DueDate
            End Get
            Set(ByVal Value As Date)
                oFinancialCommitmentInfo.DueDate = Value
            End Set
        End Property
        Public Property ThirdPartyPayment() As Boolean
            Get
                Return oFinancialCommitmentInfo.ThirdPartyPayment
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitmentInfo.ThirdPartyPayment = Value
            End Set
        End Property
        Public Property ThirdPartyPayee() As String
            Get
                Return oFinancialCommitmentInfo.ThirdPartyPayee
            End Get
            Set(ByVal Value As String)
                oFinancialCommitmentInfo.ThirdPartyPayee = Value
            End Set
        End Property
        Public Property DueDateStatement() As String
            Get
                Return oFinancialCommitmentInfo.DueDateStatement
            End Get
            Set(ByVal Value As String)
                oFinancialCommitmentInfo.DueDateStatement = Value
            End Set
        End Property
        Public Property Case_Letter() As String
            Get
                Return oFinancialCommitmentInfo.Case_Letter
            End Get
            Set(ByVal Value As String)
                oFinancialCommitmentInfo.Case_Letter = Value
            End Set
        End Property
        Public Property ERACServices() As Double
            Get
                Return oFinancialCommitmentInfo.ERACServices
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.ERACServices = Value
            End Set
        End Property
        Public Property LaboratoryServices() As Double
            Get
                Return oFinancialCommitmentInfo.LaboratoryServices

            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.LaboratoryServices = Value
            End Set
        End Property
        Public Property FixedFee() As Double
            Get
                Return oFinancialCommitmentInfo.FixedFee
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.FixedFee = Value
            End Set
        End Property
        Public Property NumberofEvents() As Integer
            Get
                Return oFinancialCommitmentInfo.NumberofEvents
            End Get
            Set(ByVal Value As Integer)
                oFinancialCommitmentInfo.NumberofEvents = Value
            End Set
        End Property
        Public Property WellAbandonment() As Double
            Get
                Return oFinancialCommitmentInfo.WellAbandonment
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.WellAbandonment = Value
            End Set
        End Property
        Public Property FreeProductRecovery() As Double
            Get
                Return oFinancialCommitmentInfo.FreeProductRecovery
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.FreeProductRecovery = Value
            End Set
        End Property
        Public Property VacuumContServices() As Double
            Get
                Return oFinancialCommitmentInfo.VacuumContServices
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.VacuumContServices = Value
            End Set
        End Property
        Public Property VacuumContServicesCnt() As Integer
            Get
                Return oFinancialCommitmentInfo.VacuumContServicesCnt
            End Get
            Set(ByVal Value As Integer)
                oFinancialCommitmentInfo.VacuumContServicesCnt = Value
            End Set
        End Property
        Public Property PTTTesting() As Double
            Get
                Return oFinancialCommitmentInfo.PTTTesting
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.PTTTesting = Value
            End Set
        End Property
        Public Property ERACVacuum() As Double
            Get
                Return oFinancialCommitmentInfo.ERACVacuum
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.ERACVacuum = Value
            End Set
        End Property
        Public Property ERACVacuumCnt() As Integer
            Get
                Return oFinancialCommitmentInfo.ERACVacuumCnt
            End Get
            Set(ByVal Value As Integer)
                oFinancialCommitmentInfo.ERACVacuumCnt = Value
            End Set
        End Property
        Public Property ERACSampling() As Double
            Get
                Return oFinancialCommitmentInfo.ERACSampling
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.ERACSampling = Value
            End Set
        End Property
        Public Property IRACServicesEstimate() As Double
            Get
                Return oFinancialCommitmentInfo.IRACServicesEstimate
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.IRACServicesEstimate = Value
            End Set
        End Property
        Public Property SubContractorSvcs() As Double
            Get
                Return oFinancialCommitmentInfo.SubContractorSvcs
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.SubContractorSvcs = Value
            End Set
        End Property
        Public Property ORCContractorSvcs() As Double
            Get
                Return oFinancialCommitmentInfo.ORCContractorSvcs
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.ORCContractorSvcs = Value
            End Set
        End Property
        Public Property REMContractorSvcs() As Double
            Get
                Return oFinancialCommitmentInfo.REMContractorSvcs
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.REMContractorSvcs = Value
            End Set
        End Property

        Public Property PreInstallSetup() As Double
            Get
                Return oFinancialCommitmentInfo.PreInstallSetup
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.PreInstallSetup = Value
            End Set
        End Property


        Public Property InstallSetup() As Double
            Get
                Return oFinancialCommitmentInfo.InstallSetup
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.InstallSetup = Value
            End Set
        End Property

        Public Property MonthlySystemUse() As Double
            Get
                Return oFinancialCommitmentInfo.MonthlySystemUse
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.MonthlySystemUse = Value
            End Set
        End Property
        Public Property MonthlySystemUseCnt() As Integer
            Get
                Return oFinancialCommitmentInfo.MonthlySystemUseCnt
            End Get
            Set(ByVal Value As Integer)
                oFinancialCommitmentInfo.MonthlySystemUseCnt = Value
            End Set
        End Property
        Public Property MonthlyOMSampling() As Double
            Get
                Return oFinancialCommitmentInfo.MonthlyOMSampling
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.MonthlyOMSampling = Value
            End Set
        End Property
        Public Property MonthlyOMSamplingCnt() As Integer
            Get
                Return oFinancialCommitmentInfo.MonthlyOMSamplingCnt
            End Get
            Set(ByVal Value As Integer)
                oFinancialCommitmentInfo.MonthlyOMSamplingCnt = Value
            End Set
        End Property
        Public Property TriAnnualOMSampling() As Double
            Get
                Return oFinancialCommitmentInfo.TriAnnualOMSampling
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.TriAnnualOMSampling = Value
            End Set
        End Property
        Public Property TriAnnualOMSamplingCnt() As Integer
            Get
                Return oFinancialCommitmentInfo.TriAnnualOMSamplingCnt
            End Get
            Set(ByVal Value As Integer)
                oFinancialCommitmentInfo.TriAnnualOMSamplingCnt = Value
            End Set
        End Property
        Public Property EstimateTriAnnualLab() As Double
            Get
                Return oFinancialCommitmentInfo.EstimateTriAnnualLab
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.EstimateTriAnnualLab = Value
            End Set
        End Property
        Public Property EstimateTriAnnualLabCnt() As Integer
            Get
                Return oFinancialCommitmentInfo.EstimateTriAnnualLabCnt
            End Get
            Set(ByVal Value As Integer)
                oFinancialCommitmentInfo.EstimateTriAnnualLabCnt = Value
            End Set
        End Property
        Public Property EstimateUtilities() As Double
            Get
                Return oFinancialCommitmentInfo.EstimateUtilities
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.EstimateUtilities = Value
            End Set
        End Property
        Public Property EstimateUtilitiesCnt() As Integer
            Get
                Return oFinancialCommitmentInfo.EstimateUtilitiesCnt
            End Get
            Set(ByVal Value As Integer)
                oFinancialCommitmentInfo.EstimateUtilitiesCnt = Value
            End Set
        End Property
        Public Property ThirdPartySettlement() As Double
            Get
                Return oFinancialCommitmentInfo.ThirdPartySettlement
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.ThirdPartySettlement = Value
            End Set
        End Property
        Public Property CostRecovery() As Double
            Get
                Return oFinancialCommitmentInfo.CostRecovery
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.CostRecovery = Value
            End Set
        End Property
        Public Property Markup() As Double
            Get
                Return oFinancialCommitmentInfo.Markup
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitmentInfo.Markup = Value
            End Set
        End Property

        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oFinancialCommitmentInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFinancialCommitmentInfo.CreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFinancialCommitmentInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oFinancialCommitmentInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitmentInfo.Deleted = Value
            End Set
        End Property


        Public Property ReimburseERAC() As Boolean
            Get
                Return oFinancialCommitmentInfo.ReimburseERAC
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitmentInfo.ReimburseERAC = Value
            End Set
        End Property




        ' The entity ID associated with a financialtext object.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return oFinancialCommitmentInfo.EntityID
            End Get
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oFinancialCommitmentInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitmentInfo.IsDirty = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFinancialCommitmentInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFinancialCommitmentInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFinancialCommitmentInfo.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "General Operations"
        Public Sub Clear()
            oFinancialCommitmentInfo = New MUSTER.Info.FinancialCommitmentInfo
        End Sub
        Public Sub Reset()
            oFinancialCommitmentInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        Public Function Retrieve(ByVal CommitmentID As Int64) As MUSTER.Info.FinancialCommitmentInfo
            Try
                oFinancialCommitmentInfo = oFinancialCommitmentDB.DBGetByID(CommitmentID)
                If oFinancialCommitmentInfo.CommitmentID = 0 Then
                    nID -= 1
                End If

                Return oFinancialCommitmentInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If Me.ValidateData(strModuleName) Then

                    oFinancialCommitmentDB.Put(oFinancialCommitmentInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFinancialCommitmentInfo.Archive()

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
            Dim xFinancialCommitmentInfo As MUSTER.Info.FinancialCommitmentInfo
            Try
                For Each xFinancialCommitmentInfo In colFinancialCommitment.Values
                    If xFinancialCommitmentInfo.IsDirty Then
                        oFinancialCommitmentInfo = xFinancialCommitmentInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oFinancialCommitmentInfoInfo As MUSTER.Info.FinancialCommitmentInfo)
            Try
                colFinancialCommitment.Add(oFinancialCommitmentInfoInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFinancialCommitmentInfo As Object)
            Try
                colFinancialCommitment.Remove(oFinancialCommitmentInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        ' Gets all the info
        Public Function GetAllByFinancialEvent(ByVal EventID As Int64) As MUSTER.Info.FinancialCommitmentCollection
            Try
                colFinancialCommitment.Clear()
                colFinancialCommitment = oFinancialCommitmentDB.DBGetByFinancialEvent(EventID)
                Return colFinancialCommitment
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function


#End Region
#Region " Populate Routines "


        'Public Function GetFinancialTextTable(ByVal TextType As Integer) As DataTable
        '    Dim dsReturn As New DataSet
        '    Dim dtReturn As DataTable
        '    Dim strSQL As String


        '    strSQL = "select * from tblSYS_Text where Reason_Type = " & TextType
        '    strSQL &= " and deleted = 0 Order By Text_Name"

        '    Try
        '        dsReturn = oFinancialCommitmentDB.DBGetDS(strSQL)
        '        If dsReturn.Tables(0).Rows.Count > 0 Then
        '            dtReturn = dsReturn.Tables(0)
        '        End If
        '        Return dtReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        'End Function
        Public Function PopulateFinancialFundingTypes() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialFundingTypes", False, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateFinancialContractTypes() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialContractTypes", False, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        '




        Public Function GetCommitmentOffset(ByVal nCommitmentID As Integer) As Single
            Dim dsReturn As New DataSet
            Dim sReturn As Single
            Dim strSQL As String
            Try
                sReturn = 0
                strSQL = "Select cast(Adjustment as money) - cast(Payment as Money) as Offset from vFinancialCommitment_Grid where CommitmentID = " & nCommitmentID.ToString
                dsReturn = oFinancialCommitmentDB.DBGetDS(strSQL)
                If dsReturn.Tables.Count > 0 Then
                    If dsReturn.Tables(0).Rows.Count > 0 Then
                        sReturn = dsReturn.Tables(0).Rows(0)("Offset")
                    End If
                End If
                Return sReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub RolloverID()
            Try
                oFinancialCommitmentInfo.CommitmentID = 0
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub

        Public Function GetApprovalDocuments(ByVal nCommitmentID As Integer) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "select document_location+document_name as docName from tblsys_document_manager where"
                strSQL += " date_edited in ("
                strSQL += "select max(date_edited) from tblsys_document_manager where entity_id=" + nCommitmentID.ToString + " and document_description in ('Financial Approval RSOM Cover Letter', 'Financial Approval Cover Letter') union "
                strSQL += "select max(date_edited) from tblsys_document_manager where entity_id=" + nCommitmentID.ToString + " and document_description in ('Financial Approval Form') union "
                strSQL += "select max(date_edited) from tblsys_document_manager where entity_id=" + nCommitmentID.ToString + " and document_description in ('Financial Commitment Memo')) "
                strSQL += "and entity_id=" + nCommitmentID.ToString
                dsReturn = oFinancialCommitmentDB.DBGetDS(strSQL)
                If dsReturn.Tables.Count > 0 Then
                    If dsReturn.Tables(0).Rows.Count > 0 Then
                        dtReturn = dsReturn.Tables(0)
                    End If
                End If
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
                dsReturn = oFinancialCommitmentDB.DBGetDS(strSQL)
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
            Dim strArr() As String = colFinancialCommitment.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.CommitmentID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colFinancialCommitment.Item(nArr.GetValue(colIndex + direction)).CommitmentID.ToString
            Else
                Return colFinancialCommitment.Item(nArr.GetValue(colIndex)).CommitmentID.ToString
            End If
        End Function
#End Region

    End Class

End Namespace