
Namespace MUSTER.BusinessLogic
    '  -------------------------------------------------------------------------------
    '  MUSTER.BusinessLogic.pFeeBasis
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
    Public Class pFeeBasis

#Region "Public Events"
        Public Event FeeBasisBLChanged As FeeBasisBLChangedEventHandler
        Public Event FeeBasisBLColChanged As FeeBasisBLColChangedEventHandler
        Public Event FeeBasisBLErr As FeeBasisBLErrEventHandler
        Public Event FeeBasisInfChanged As FeeBasisInfoChanged

        Public Delegate Sub FeeBasisBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeBasisBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeBasisBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FeeBasisInfoChanged()
#End Region
#Region "Private member variables"
        Private WithEvents oFeesBasis As MUSTER.Info.FeeBasisInfo
        Private WithEvents oFeesBasisCol As New MUSTER.Info.FeeBasisCollection
        Private oFeesBasisDB As New MUSTER.DataAccess.FeeBasisDB

        Private MusterExceptions As Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            oFeesBasisDB = New MUSTER.dataaccess.FeeBasisDB
            oFeesBasisCol = New MUSTER.Info.FeeBasisCollection
        End Sub
        Public Sub New(ByVal strName As String)
            oFeesBasisDB = New MUSTER.dataaccess.FeeBasisDB
            oFeesBasisCol = New MUSTER.Info.FeeBasisCollection
        End Sub
        Public Sub New(ByVal FeeBasisID As Integer)
            oFeesBasisDB = New MUSTER.dataaccess.FeeBasisDB
            oFeesBasisCol = New MUSTER.Info.FeeBasisCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        ' The base fee for the billing period
        Public Property BaseFee() As Decimal
            Get
                Return oFeesBasis.BaseFee
            End Get
            Set(ByVal Value As Decimal)
                oFeesBasis.BaseFee = Value
            End Set
        End Property
        ' The billing unit for the base fee (from tblSYS_PROPERTY)
        Public Property BaseUnit() As Long
            Get
                Return oFeesBasis.BaseUnit
            End Get
            Set(ByVal Value As Long)
                oFeesBasis.BaseUnit = Value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
            End Get
            Set(ByVal Value As Boolean)
            End Set
        End Property
        ' Gets/Sets the active flag for the financial text info object (from info.Active)
        Public Property Completed() As Boolean
            Get

            End Get
            Set(ByVal Value As Boolean)
            End Set
        End Property
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oFeesBasis.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFeesBasis.CreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFeesBasis.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oFeesBasis.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFeesBasis.Deleted = Value
            End Set
        End Property
        ' The description associated with the Fee Basis.
        Public Property Description() As String
            Get
                Return oFeesBasis.Description
            End Get
            Set(ByVal Value As String)
                oFeesBasis.Description = Value
            End Set
        End Property
        ' The early grace date for the billing period
        Public Property EarlyGrace() As Date
            Get
                Return oFeesBasis.EarlyGrace
            End Get
            Set(ByVal Value As Date)
                oFeesBasis.EarlyGrace = Value
            End Set
        End Property
        ' The entity ID associated with a financialtext object.
        Public ReadOnly Property EntityID() As Integer
            Get
            End Get
        End Property



        ' The fiscal year for the billing period
        Public Property FiscalYear() As Int32
            Get
                Return oFeesBasis.FiscalYear
            End Get
            Set(ByVal Value As Int32)
                oFeesBasis.FiscalYear = Value
            End Set
        End Property
        ' The unique ID for the row containing the text in the table (info.ID)
        Public ReadOnly Property ID() As Int64
            Get
                Return oFeesBasis.ID
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oFeesBasis.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFeesBasis.IsDirty = Value
            End Set
        End Property
        ' The base late fee for the billing period
        Public Property LateFee() As Decimal
            Get
                Return oFeesBasis.LateFee
            End Get
            Set(ByVal Value As Decimal)
                oFeesBasis.LateFee = Value
            End Set
        End Property
        ' The late grace date for the billing period
        Public Property LateGrace() As Date
            Get
                Return oFeesBasis.LateGrace
            End Get
            Set(ByVal Value As Date)
                oFeesBasis.LateGrace = Value
            End Set
        End Property
        ' ??? Always a period of time (from tblSYS_PROPERTY_MASTER)
        Public Property LatePeriod() As Long
            Get
                Return oFeesBasis.LatePeriod
            End Get
            Set(ByVal Value As Long)
                oFeesBasis.LatePeriod = Value
            End Set
        End Property
        ' The type of late fee calculation involving the LateFee for the billing period (from tblSYS_PROPERTY_MASTER)
        Public Property LateType() As Long
            Get
                Return oFeesBasis.LateType
            End Get
            Set(ByVal Value As Long)
                oFeesBasis.LateType = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFeesBasis.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFeesBasis.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFeesBasis.ModifiedOn
            End Get
        End Property
        ' The last date for the billing period
        Public Property PeriodEnd() As Date
            Get
                Return oFeesBasis.PeriodEnd
            End Get
            Set(ByVal Value As Date)
                oFeesBasis.PeriodEnd = Value
            End Set
        End Property
        ' The start date for the billing period
        Public Property PeriodStart() As Date
            Get
                Return oFeesBasis.PeriodStart
            End Get
            Set(ByVal Value As Date)
                oFeesBasis.PeriodStart = Value
            End Set
        End Property

        Public Property GenerateDate() As Date
            Get
                Return oFeesBasis.GenerateDate
            End Get
            Set(ByVal Value As Date)
                oFeesBasis.GenerateDate = Value
            End Set
        End Property
        Public Property Generated() As Boolean
            Get
                Return oFeesBasis.Generated
            End Get
            Set(ByVal Value As Boolean)
                oFeesBasis.Generated = Value
            End Set
        End Property
        Public Property GenerateTime() As Date
            Get
                Return oFeesBasis.GenerateTime
            End Get
            Set(ByVal Value As Date)
                oFeesBasis.GenerateTime = Value
            End Set
        End Property

        Public Property ApprovedTime() As Date
            Get
                Return oFeesBasis.ApprovedTime
            End Get
            Set(ByVal Value As Date)
                oFeesBasis.ApprovedTime = Value
            End Set
        End Property
        Public Property ApprovedDate() As Date
            Get
                Return oFeesBasis.ApprovedDate
            End Get
            Set(ByVal Value As Date)
                oFeesBasis.ApprovedDate = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal FeeBasisID As Int64) As MUSTER.Info.FeeBasisInfo
            Dim oFeesBasisInfoLocal As MUSTER.Info.FeeBasisInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oFeesBasisInfoLocal In oFeesBasisCol.Values
                    If oFeesBasisInfoLocal.ID = ID Then
                        If oFeesBasisInfoLocal.IsAgedData = True And oFeesBasisInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oFeesBasis = oFeesBasisInfoLocal
                            Return oFeesBasis
                        End If
                    End If
                Next
                If bolDataAged Then
                    oFeesBasisCol.Remove(oFeesBasisInfoLocal)
                End If
                oFeesBasis = oFeesBasisDB.DBGetByID(FeeBasisID)
                oFeesBasisCol.Add(oFeesBasis)
                Return oFeesBasis
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Obtains and returns an entity as called for by ID
        Public Function GetByFiscalYear(ByVal FeeBasisYear As Int32) As MUSTER.Info.FeeBasisInfo
            Dim oFeesBasisInfoLocal As MUSTER.Info.FeeBasisInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oFeesBasisInfoLocal In oFeesBasisCol.Values
                    If oFeesBasisInfoLocal.FiscalYear = FeeBasisYear Then
                        If oFeesBasisInfoLocal.IsAgedData = True And oFeesBasisInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oFeesBasis = oFeesBasisInfoLocal
                            Return oFeesBasis
                        End If
                    End If
                Next
                If bolDataAged Then
                    oFeesBasisCol.Remove(oFeesBasisInfoLocal)
                End If
                oFeesBasis = oFeesBasisDB.DBGetByFiscalYear(FeeBasisYear)
                oFeesBasisCol.Add(oFeesBasis)
                Return oFeesBasis
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Obtains and returns an entity as called for by ID
        Public Function GetFiscalYear(ByVal feeBasisDate As DateTime) As Int16
            Dim FiscalYear As Int16
            Try

                FiscalYear = oFeesBasisDB.DBGetFiscalYear(FeeBasisdate)

                Return FiscalYear
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetFiscalYearForFee(Optional ByVal feeBasisDate As DateTime = #1/1/1900#) As Int16
            Dim FiscalYear As Int16
            Try

                FiscalYear = oFeesBasisDB.DBGetFiscalYearForFees(feeBasisDate)

                Return FiscalYear
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
                'If Me.GenerateDate < Now.Date Then
                '    validateSuccess = False
                'End If
                'If Me.ApprovedDate <> "01/01/0001" And Me.ApprovedDate < Me.GenerateDate And Me.ID > 0 Then
                '    validateSuccess = False
                'End If
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
                    oFeesBasisDB.put(oFeesBasis, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFeesBasis.IsDirty = False
                    oFeesBasis.Archive()
                    RaiseEvent FeeBasisBLChanged(oFeesBasis.IsDirty)
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
                ''oFeesBasis = oFeesBasisDB.DBGetByID(ID)
                oFeesBasis.ID = ID
                If oFeesBasis.ID = 0 Then
                    'oFeesBasis.ID = nID
                    'nID -= 1
                End If
                oFeesBasisCol.Add(oFeesBasis)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef FeesBasis As MUSTER.Info.FeeBasisInfo)
            Try
                oFeesBasis = FeesBasis
                'oFeesBasis.UserID = onUserID
                oFeesBasisCol.Add(oFeesBasis)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oFeesBasisLocal As MUSTER.Info.FeeBasisInfo

            Try
                For Each oFeesBasisLocal In oFeesBasisCol.Values
                    If oFeesBasisLocal.ID = ID Then
                        oFeesBasisCol.Remove(oFeesBasisLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Pipe " & ID.ToString & " is not in the collection of Pipes.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oLustEvent As MUSTER.Info.FeeBasisInfo)
            Try
                oFeesBasisCol.Remove(oLustEvent)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LustEvent " & oLustEvent.ID & " is not in the collection of FeeBasis.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xFeesBasisInfo As MUSTER.Info.FeeBasisInfo
            For Each xFeesBasisInfo In oFeesBasisCol.Values
                If xFeesBasisInfo.IsDirty Then
                    oFeesBasis = xFeesBasisInfo
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
                    If oFeesBasis.ID <= 0 Then
                        oFeesBasis.CreatedBy = UserID
                    Else
                        oFeesBasis.ModifiedBy = UserID
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
            '    Dim strArr() As String = oFeesBasisCol.GetKeys()
            '    Dim nArr(strArr.GetUpperBound(0)) As Integer
            '    Dim y As String
            '    For Each y In strArr
            '        nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            '    Next
            '    nArr.Sort(nArr)
            '    colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            '    If colIndex + direction > -1 And _
            '        colIndex + direction <= nArr.GetUpperBound(0) Then
            '        Return oFeesBasisCol.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            '    Else
            '        Return oFeesBasisCol.Item(nArr.GetValue(colIndex)).ID.ToString
            '    End If
        End Function
#End Region
#Region "General Operations"

        Public Function Reset() As Object
            oFeesBasis.Reset()
            oFeesBasisCol.Clear()
        End Function
#End Region
#Region "Miscellaneous Operations"

#End Region

#Region " Populate Routines "
        Public Function ApproveBilling() As String
            Dim strReturn As String
            Dim strSQL As String
            Try
                strSQL = "UPDATE dbo.tblFEES_PENDING_INVOICES SET PROCESSED = 1 "
                oFeesBasisDB.DBExecNonQuery(strSQL)


                strSQL = "EXEC spFeesInsertToPermInvBP2KTemp"
                oFeesBasisDB.DBExecNonQuery(strSQL)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function MarkCalendarCompleted_ByDesc(ByVal strDesc As String) As String
            Dim strReturn As String
            Dim strSQL As String
            Try
                strSQL = "UPDATE dbo.tblSYS_CALENDAR_CALENDAR_INFO SET COMPLETED = 1 WHERE Owning_Entity_Type = 25 AND Task_Description like '%" & strDesc & "%'"
                oFeesBasisDB.DBExecNonQuery(strSQL)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PurgePendingInvoiceHeaders() As Boolean
            Dim strSQL As String
            Try
                strSQL = "Delete from tblFEES_PENDING_INVOICES where FEE_TYPE = 'FD'"

                oFeesBasisDB.DBExecNonQuery(strSQL)

                Return True

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetFeesPendingInvoiceHeaders() As DataSet
            Dim dsReturn As New DataSet
            Dim strSQL As String

            strSQL = ""
            strSQL &= " SELECT * FROM vFeesPendingInvoiceHeaders"

            Try
                dsReturn = oFeesBasisDB.DBGetDS(strSQL)

                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetFeesBasisGrid() As DataSet
            Dim dsReturn As New DataSet
            Dim strSQL As String

            strSQL = ""
            strSQL &= " SELECT * FROM vFeesBasisGrid ORDER BY FISCAL_YEAR DESC"

            Try
                dsReturn = oFeesBasisDB.DBGetDS(strSQL)

                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetMaxFeeBasisID() As Int64
            Dim dsRemSys As New DataSet
            Dim dtReturn As Int64
            Dim strSQL As String

            Try

                strSQL = "select  * "
                strSQL &= "from vFeesBasisGrid "
                strSQL &= "order by Fiscal_Year Desc, Fees_Basis_ID DESC "

                dsRemSys = oFeesBasisDB.DBGetDS(strSQL)
                dtReturn = 0
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0).Rows(0)("Fees_Basis_ID")
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateFeeUnits() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFeeUnitsList")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateLateFeePeriod() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFeeLatePeriodList")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateLateFeeType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFeeLateTypeList")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
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
                dsReturn = oFeesBasisDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
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

