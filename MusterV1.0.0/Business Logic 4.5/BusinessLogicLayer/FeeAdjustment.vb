

Namespace MUSTER.BusinessLogic
    '  -------------------------------------------------------------------------------
    '  MUSTER.BusinessLogic.pFeeAdjustment
    '  Provides the operations required to manipulate a Fee Basis object.
    '  
    '  Copyright (C) 2004, 2005 CIBER, Inc.
    '  All rights reserved.
    '  
    '  Release   Initials    Date        Description
    '  1.0         JC       06/14/2005    Original class definition
    '  
    '  Function          Description
    '  -------------------------------------------------------------------------------
    '  Attribute          Description
    '  -------------------------------------------------------------------------------
    Public Class pFeeAdjustment

#Region "Public Events"
        Public Event FeeAdjustmentBLChanged As FeeAdjustmentBLChangedEventHandler
        Public Event FeeAdjustmentBLColChanged As FeeAdjustmentBLColChangedEventHandler
        Public Event FeeAdjustmentBLErr As FeeAdjustmentBLErrEventHandler
        Public Event FeeAdjustmentInfChanged As FeeAdjustmentInfoChanged

        Public Delegate Sub FeeAdjustmentBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeAdjustmentBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeAdjustmentBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FeeAdjustmentInfoChanged()
#End Region
#Region "Private member variables"
        Private WithEvents oFeeAdjustment As MUSTER.Info.FeeAdjustmentInfo
        Private WithEvents oFeeAdjustmentCol As New MUSTER.Info.FeeAdjustmentCollection
        Private oFeeAdjustmentDB As New MUSTER.DataAccess.FeeAdjustmentDB

        Private MusterExceptions As Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            oFeeAdjustmentDB = New MUSTER.dataaccess.FeeAdjustmentDB
            oFeeAdjustmentCol = New MUSTER.Info.FeeAdjustmentCollection
        End Sub
        Public Sub New(ByVal strName As String)
            oFeeAdjustmentDB = New MUSTER.dataaccess.FeeAdjustmentDB
            oFeeAdjustmentCol = New MUSTER.Info.FeeAdjustmentCollection
        End Sub
        Public Sub New(ByVal FeeAdjustmentID As Integer)
            oFeeAdjustmentDB = New MUSTER.dataaccess.FeeAdjustmentDB
            oFeeAdjustmentCol = New MUSTER.Info.FeeAdjustmentCollection
        End Sub
#End Region
#Region "Exposed Attributes"



        '                                                        AltIsDBNull(drSet.Item("Check_Number"), String.Empty), _
        '                                                        AltIsDBNull(drSet.Item("Reason"), String.Empty), _
        '                                                        AltIsDBNull(drSet.Item("Returned_From_BP2K"), 0))


        Public Property ID() As Decimal
            Get
                Return oFeeAdjustment.ID
            End Get
            Set(ByVal Value As Decimal)
                oFeeAdjustment.ID = Value
            End Set
        End Property

        Public Property OwnerID() As Int64
            Get
                Return oFeeAdjustment.OwnerID
            End Get
            Set(ByVal Value As Int64)
                oFeeAdjustment.OwnerID = Value
            End Set
        End Property

        Public Property Deleted() As Boolean
            Get
                Return oFeeAdjustment.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFeeAdjustment.Deleted = Value
            End Set
        End Property

        Public Property CreditCode() As String
            Get
                Return oFeeAdjustment.CreditCode
            End Get
            Set(ByVal Value As String)
                oFeeAdjustment.CreditCode = Value
            End Set
        End Property

        Public Property FiscalYear() As Int32
            Get
                Return oFeeAdjustment.FiscalYear
            End Get
            Set(ByVal Value As Int32)
                oFeeAdjustment.FiscalYear = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oFeeAdjustment.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFeeAdjustment.IsDirty = Value
            End Set
        End Property

        Public Property Amount() As Decimal
            Get
                Return oFeeAdjustment.Amount
            End Get
            Set(ByVal Value As Decimal)
                oFeeAdjustment.Amount = Value
            End Set
        End Property

        Public Property Applied() As Date
            Get
                Return oFeeAdjustment.Applied
            End Get
            Set(ByVal Value As Date)
                oFeeAdjustment.Applied = Value
            End Set
        End Property

        Public Property FacilityID() As Int64
            Get
                Return oFeeAdjustment.FacilityID
            End Get
            Set(ByVal Value As Int64)
                oFeeAdjustment.FacilityID = Value
            End Set
        End Property

        Public Property InvoiceNumber() As String
            Get
                Return oFeeAdjustment.InvoiceNumber
            End Get
            Set(ByVal Value As String)
                oFeeAdjustment.InvoiceNumber = Value
            End Set
        End Property

        Public Property ItemSeqNumber() As String
            Get
                Return oFeeAdjustment.ItemSeqNumber
            End Get
            Set(ByVal Value As String)
                oFeeAdjustment.ItemSeqNumber = Value
            End Set
        End Property

        Public Property CheckNumber() As String
            Get
                Return oFeeAdjustment.CheckNumber
            End Get
            Set(ByVal Value As String)
                oFeeAdjustment.CheckNumber = Value
            End Set
        End Property

        Public Property Reason() As String
            Get
                Return oFeeAdjustment.Reason
            End Get
            Set(ByVal Value As String)
                oFeeAdjustment.Reason = Value
            End Set
        End Property

        Public ReadOnly Property ReturnedFromBP2K() As Boolean
            Get
                Return oFeeAdjustment.ReturnedFromBP2K
            End Get
        End Property

        Public Property CreatedBy() As String
            Get
                Return oFeeAdjustment.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFeeAdjustment.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFeeAdjustment.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFeeAdjustment.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFeeAdjustment.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFeeAdjustment.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal FeeAdjustmentID As Int64) As MUSTER.Info.FeeAdjustmentInfo
            Dim oFeeAdjustmentInfoLocal As MUSTER.Info.FeeAdjustmentInfo
            Dim bolDataAged As Boolean = False
            Try
                If FeeAdjustmentID = 0 Then
                    oFeeAdjustment = New MUSTER.Info.FeeAdjustmentInfo
                    Return oFeeAdjustment
                End If

                For Each oFeeAdjustmentInfoLocal In oFeeAdjustmentCol.Values
                    If oFeeAdjustmentInfoLocal.ID = ID Then
                        If oFeeAdjustmentInfoLocal.IsAgedData = True And oFeeAdjustmentInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oFeeAdjustment = oFeeAdjustmentInfoLocal
                            Return oFeeAdjustment
                        End If
                    End If
                Next
                If bolDataAged Then
                    oFeeAdjustmentCol.Remove(oFeeAdjustmentInfoLocal)
                End If
                oFeeAdjustment = oFeeAdjustmentDB.DBGetByID(FeeAdjustmentID)
                oFeeAdjustmentCol.Add(oFeeAdjustment)
                Return oFeeAdjustment
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function RetrieveOwnerOverages(ByVal OwnerID As Int64) As DataSet
            Dim dsReturn As New DataSet
            Dim dsRel1 As DataRelation
            Dim strSQL As String


            strSQL = "select  SFY, OWNER_ID, convert(nvarchar(100),FACILITY_ID) as FACILITY_ID, INV_NUMBER, convert(varchar, convert(decimal(9, 2), INV_AMT),1)  as INV_AMT, DATE_APPLIED, CHECK_NUMBER, BP2K_Trans_ID, REASON, 'AP' as TYPE from dbo.tblFEES_ADJUSTMENTS " + _
                      " where Deleted = 0 And owner_id = " + OwnerID.ToString + _
                      " union   select SFY, OWNER_ID, 'N/A' as FACILITY_ID, isnull(INV_NUMBER,'Owner Level') as INV_NUMBER, convert(varchar,isnull(AMT_RECEIVED,0.0),1) as INV_AMT, RECEIPT_DATE as DATE_APPLIED, CHECK_NUMBER, BP2K_Trans_ID, case when MISAPPLY_FLAg like '%RA%' then 'Re-Apply Oeverage' else 'Mis-Applied Overage' end AS REASON, MISAPPLY_FLAG as Type from tblFEES_RECEIPTS " + _
                      " where deleted = 0 and owner_id = " + OwnerID.ToString + "  and ((inv_number is null and MISAPPLY_FLAG = 'DR') or MISAPPLy_FLAG = 'RA') order by date_applied, type desc"
            Try
                dsReturn = oFeeAdjustmentDB.DBGetDS(strSQL)

               Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True

            Try

                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If ValidateData() Then
                    oFeeAdjustmentDB.put(oFeeAdjustment, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFeeAdjustment.IsDirty = False
                    oFeeAdjustment.Archive()
                    RaiseEvent FeeAdjustmentBLChanged(oFeeAdjustment.IsDirty)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                ''oFeeAdjustment = oFeeAdjustmentDB.DBGetByID(ID)
                oFeeAdjustment.ID = ID
                If oFeeAdjustment.ID = 0 Then
                    'oFeeAdjustment.ID = nID
                    'nID -= 1
                End If
                oFeeAdjustmentCol.Add(oFeeAdjustment)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef FeeAdjustment As MUSTER.Info.FeeAdjustmentInfo)
            Try
                oFeeAdjustment = FeeAdjustment
                'oFeeAdjustment.UserID = onUserID
                oFeeAdjustmentCol.Add(oFeeAdjustment)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oFeeAdjustmentLocal As MUSTER.Info.FeeAdjustmentInfo

            Try
                For Each oFeeAdjustmentLocal In oFeeAdjustmentCol.Values
                    If oFeeAdjustmentLocal.ID = ID Then
                        oFeeAdjustmentCol.Remove(oFeeAdjustmentLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Fee Adjustment " & ID.ToString & " is not in the collection of Fee Adjustments.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFeeAdjustment As MUSTER.Info.FeeAdjustmentInfo)
            Try
                oFeeAdjustmentCol.Remove(oFeeAdjustment)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Fee Adjustment " & oFeeAdjustment.ID & " is not in the collection of oFeeAdjustment.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xFeeAdjustmentInfo As MUSTER.Info.FeeAdjustmentInfo
            For Each xFeeAdjustmentInfo In oFeeAdjustmentCol.Values
                If xFeeAdjustmentInfo.IsDirty Then
                    oFeeAdjustment = xFeeAdjustmentInfo
                    '********************************************************
                    '
                    ' Note that if there are contained objects and the respective
                    '  contained object collections are dirty, then the contained
                    '  collections MUST BE FLUSHED before this object can be
                    '  saved.  Otherwise, there is a risk that an attempt will
                    '  be made to insert a new object to the repository without
                    '  corresponding contained information being present which 
                    '  may, in turn, cause a foreign key violation!
                    '
                    '********************************************************
                    If oFeeAdjustment.ID <= 0 Then
                        oFeeAdjustment.CreatedBy = UserID
                    Else
                        oFeeAdjustment.ModifiedBy = UserID
                    End If
                    Me.Save(moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                End If
            Next
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            '    Dim strArr() As String = oFeeAdjustmentCol.GetKeys()
            '    Dim nArr(strArr.GetUpperBound(0)) As Integer
            '    Dim y As String
            '    For Each y In strArr
            '        nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            '    Next
            '    nArr.Sort(nArr)
            '    colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            '    If colIndex + direction > -1 And _
            '        colIndex + direction <= nArr.GetUpperBound(0) Then
            '        Return oFeeAdjustmentCol.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            '    Else
            '        Return oFeeAdjustmentCol.Item(nArr.GetValue(colIndex)).ID.ToString
            '    End If
        End Function
#End Region
#Region "General Operations"

        Public Function Reset() As Object
            oFeeAdjustment.Reset()
            oFeeAdjustmentCol.Clear()
        End Function
#End Region
#Region "Miscellaneous Operations"

#End Region

#Region " Populate Routines "
        'Public Function ApproveBilling() As String
        '    Dim strReturn As String
        '    Dim strSQL As String
        '    Try
        '        strSQL = "UPDATE dbo.tblFEES_PENDING_INVOICES SET PROCESSED = 1 "
        '        oFeeAdjustmentDB.DBExecNonQuery(strSQL)


        '        strSQL = "EXEC spFeesInsertToPermInvBP2KTemp"
        '        oFeeAdjustmentDB.DBExecNonQuery(strSQL)

        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        'Public Function MarkCalendarCompleted_ByDesc(ByVal strDesc As String) As String
        '    Dim strReturn As String
        '    Dim strSQL As String
        '    Try
        '        strSQL = "UPDATE dbo.tblSYS_CALENDAR_CALENDAR_INFO SET COMPLETED = 1 WHERE Owning_Entity_Type = 25 AND Task_Description like '%" & strDesc & "%'"
        '        oFeeAdjustmentDB.DBExecNonQuery(strSQL)

        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        'Public Function PurgePendingInvoiceHeaders() As Boolean
        '    Dim strSQL As String

        '    strSQL = ""
        '    strSQL &= "Delete from tblFEES_PENDING_INVOICES where FEE_TYPE = 'FD' and Processed = 0"

        '    Try
        '        Return oFeeAdjustmentDB.DBExecNonQuery(strSQL)

        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        'Public Function GetFeesPendingInvoiceHeaders() As DataSet
        '    Dim dsReturn As New DataSet
        '    Dim strSQL As String

        '    strSQL = ""
        '    strSQL &= " SELECT * FROM vFeesPendingInvoiceHeaders"

        '    Try
        '        dsReturn = oFeeAdjustmentDB.DBGetDS(strSQL)

        '        Return dsReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        'Public Function GetFeeAdjustmentGrid() As DataSet
        '    Dim dsReturn As New DataSet
        '    Dim strSQL As String

        '    strSQL = ""
        '    strSQL &= " SELECT * FROM vFeeAdjustmentGrid"

        '    Try
        '        dsReturn = oFeeAdjustmentDB.DBGetDS(strSQL)

        '        Return dsReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function




        'Public Function PopulateFeeUnits() As DataTable
        '    Try
        '        Dim dtReturn As DataTable = GetDataTable("vFeeUnitsList")
        '        Return dtReturn
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function

        'Public Function PopulateLateFeePeriod() As DataTable
        '    Try
        '        Dim dtReturn As DataTable = GetDataTable("vFeeLatePeriodList")
        '        Return dtReturn
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function

        'Public Function PopulateLateFeeType() As DataTable
        '    Try
        '        Dim dtReturn As DataTable = GetDataTable("vFeeLateTypeList")
        '        Return dtReturn
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        'Public Function PopulateFinancialStatus() As DataTable
        '    Try
        '        Dim dtReturn As DataTable = GetDataTable("vFinancialEventStatus", false, false)
        '        Return dtReturn
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function


        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal PropertyMaster As Boolean = True, Optional ByVal IncludeBlank As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As New DataTable   'Added New
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
                dsReturn = oFeeAdjustmentDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                'Return dtReturn    'Commented SS And Added below line
                GetDataTable = dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

#End Region



#End Region
#Region "External Event Handlers"

#End Region
    End Class
End Namespace