
Namespace MUSTER.BusinessLogic
    '  -------------------------------------------------------------------------------
    '  MUSTER.BusinessLogic.pFeeLateFee
    '  Provides the operations required to manipulate a Late Fee object.
    '  
    '  Copyright (C) 2004, 2005 CIBER, Inc.
    '  All rights reserved.
    '  
    '  Release   Initials    Date        Description
    '  1.0         AB       12/05/2005    Original class definition
    '  
    '  Function          Description
    '  -------------------------------------------------------------------------------
    '  Attribute          Description
    '  -------------------------------------------------------------------------------
    Public Class pFeeLateFee

#Region "Public Events"
        Public Event FeeLateFeeBLChanged As FeeLateFeeBLChangedEventHandler
        Public Event FeeLateFeeBLColChanged As FeeLateFeeBLColChangedEventHandler
        Public Event FeeLateFeeBLErr As FeeLateFeeBLErrEventHandler
        Public Event FeeLateFeeInfChanged As FeeLateFeeInfoChanged

        Public Delegate Sub FeeLateFeeBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeLateFeeBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeLateFeeBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FeeLateFeeInfoChanged()
#End Region
#Region "Private member variables"
        Private WithEvents oFeesLateFee As MUSTER.Info.FeeLateFeeInfo
        Private WithEvents oFeesLateFeeCol As New MUSTER.Info.FeeLateFeeCollection
        Private oFeesLateFeeDB As New MUSTER.DataAccess.FeeLateFeeDB

        Private MusterExceptions As Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            oFeesLateFeeDB = New MUSTER.dataaccess.FeeLateFeeDB
            oFeesLateFeeCol = New MUSTER.Info.FeeLateFeeCollection
        End Sub
        Public Sub New(ByVal strName As String)
            oFeesLateFeeDB = New MUSTER.dataaccess.FeeLateFeeDB
            oFeesLateFeeCol = New MUSTER.Info.FeeLateFeeCollection
        End Sub
        Public Sub New(ByVal FeeLateFeeID As Integer)
            oFeesLateFeeDB = New MUSTER.dataaccess.FeeLateFeeDB
            oFeesLateFeeCol = New MUSTER.Info.FeeLateFeeCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        ' The base fee for the billing period
        Public Property CertLetterNumber() As String
            Get
                Return oFeesLateFee.CertLetterNumber
            End Get
            Set(ByVal Value As String)
                oFeesLateFee.CertLetterNumber = Value
            End Set
        End Property
        ' The billing unit for the base fee (from tblSYS_PROPERTY)
        Public Property FiscalYear() As Integer
            Get
                Return oFeesLateFee.FiscalYear
            End Get
            Set(ByVal Value As Integer)
                oFeesLateFee.FiscalYear = Value
            End Set
        End Property
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oFeesLateFee.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFeesLateFee.CreatedBy = Value
            End Set
        End Property
        Public Property WaiverFinalizedOn() As Date
            Get
                Return oFeesLateFee.WaiverFinalizedOn
            End Get
            Set(ByVal Value As Date)
                oFeesLateFee.WaiverFinalizedOn = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFeesLateFee.CreateDate
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oFeesLateFee.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFeesLateFee.Deleted = Value
            End Set
        End Property
        ' The description associated with the Fee Basis.
        Public Property InvoiceNumber() As String
            Get
                Return oFeesLateFee.InvoiceNumber
            End Get
            Set(ByVal Value As String)
                oFeesLateFee.InvoiceNumber = Value
            End Set
        End Property

        ' The unique ID for the row containing the text in the table (info.ID)
        Public ReadOnly Property ID() As Int64
            Get
                Return oFeesLateFee.ID
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oFeesLateFee.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFeesLateFee.IsDirty = Value
            End Set
        End Property

        Public Property LateCharges() As Decimal
            Get
                Return oFeesLateFee.LateCharges
            End Get
            Set(ByVal Value As Decimal)
                oFeesLateFee.LateCharges = Value
            End Set
        End Property

        Public Property ProcessCertification() As Boolean
            Get
                Return oFeesLateFee.ProcessCertification
            End Get
            Set(ByVal Value As Boolean)
                oFeesLateFee.ProcessCertification = Value
            End Set
        End Property

        Public Property ProcessWaiver() As Boolean
            Get
                Return oFeesLateFee.ProcessWaiver
            End Get
            Set(ByVal Value As Boolean)
                oFeesLateFee.ProcessWaiver = Value
            End Set
        End Property

        Public Property WaiveReason() As Long
            Get
                Return oFeesLateFee.WaiveReason
            End Get
            Set(ByVal Value As Long)
                oFeesLateFee.WaiveReason = Value
            End Set
        End Property
        ' True = Approve  -+- False = Deny
        Public Property WaiveApprovalRecommendation() As Boolean
            Get
                Return oFeesLateFee.WaiveApprovalRecommendation
            End Get
            Set(ByVal Value As Boolean)
                oFeesLateFee.WaiveApprovalRecommendation = Value
            End Set
        End Property
        ' True = Approve  -+- False = Deny
        Public Property WaiveApprovalStatus() As Boolean
            Get
                Return oFeesLateFee.WaiveApprovalStatus
            End Get
            Set(ByVal Value As Boolean)
                oFeesLateFee.WaiveApprovalStatus = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFeesLateFee.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFeesLateFee.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFeesLateFee.ModifiedDate
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal FeeLateFeeCertID As Int64) As MUSTER.Info.FeeLateFeeInfo
            Dim oFeesLateFeeInfoLocal As MUSTER.Info.FeeLateFeeInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oFeesLateFeeInfoLocal In oFeesLateFeeCol.Values
                    If oFeesLateFeeInfoLocal.ID = FeeLateFeeCertID Then
                        If oFeesLateFeeInfoLocal.IsAgedData = True And oFeesLateFeeInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oFeesLateFee = oFeesLateFeeInfoLocal
                            Return oFeesLateFee
                        End If
                    End If
                Next
                If bolDataAged Then
                    oFeesLateFeeCol.Remove(oFeesLateFeeInfoLocal)
                End If
                oFeesLateFee = oFeesLateFeeDB.DBGetByid(FeeLateFeeCertID)
                oFeesLateFeeCol.Add(oFeesLateFee)
                Return oFeesLateFee
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function Retrieve_ByInvoiceNumber(ByVal InvoiceNumber As String) As MUSTER.Info.FeeLateFeeInfo
            Dim oFeesLateFeeInfoLocal As MUSTER.Info.FeeLateFeeInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oFeesLateFeeInfoLocal In oFeesLateFeeCol.Values
                    If oFeesLateFeeInfoLocal.InvoiceNumber = InvoiceNumber Then
                        If oFeesLateFeeInfoLocal.IsAgedData = True And oFeesLateFeeInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oFeesLateFee = oFeesLateFeeInfoLocal
                            Return oFeesLateFee
                        End If
                    End If
                Next
                If bolDataAged Then
                    oFeesLateFeeCol.Remove(oFeesLateFeeInfoLocal)
                End If
                oFeesLateFee = oFeesLateFeeDB.DBGetByInvoiceNumber(InvoiceNumber)
                oFeesLateFeeCol.Add(oFeesLateFee)
                Return oFeesLateFee
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

        Public Function SaveRegTag(ByVal moduleID As Integer, ByVal staffID As Integer, ByVal ownerID As Integer, ByVal processed As Boolean, ByVal year As Integer, ByVal certNumber As String, ByVal facilityID As Integer, Optional ByVal ID As Integer = -1) As Integer
            Try
                Return oFeesLateFeeDB.putRegtag(moduleID, staffID, ownerID, processed, year, certNumber, facilityID, ID)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If ValidateData() Then
                    oFeesLateFeeDB.put(oFeesLateFee, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFeesLateFee.IsDirty = False
                    oFeesLateFee.Archive()
                    RaiseEvent FeeLateFeeBLChanged(oFeesLateFee.IsDirty)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Function SaveExcuse(ByVal excuse As String, ByVal userID As String) As Integer
            Try
                Return oFeesLateFeeDB.putLatFees(excuse, userID)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                ''oFeesLateFee = oFeesLateFeeDB.DBGetByID(ID)
                oFeesLateFee.ID = ID
                If oFeesLateFee.ID = 0 Then
                    'oFeesLateFee.ID = nID
                    'nID -= 1
                End If
                oFeesLateFeeCol.Add(oFeesLateFee)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef FeesLateFee As MUSTER.Info.FeeLateFeeInfo)
            Try
                oFeesLateFee = FeesLateFee
                'oFeesLateFee.UserID = onUserID
                oFeesLateFeeCol.Add(oFeesLateFee)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oFeesLateFeeLocal As MUSTER.Info.FeeLateFeeInfo

            Try
                For Each oFeesLateFeeLocal In oFeesLateFeeCol.Values
                    If oFeesLateFeeLocal.ID = ID Then
                        oFeesLateFeeCol.Remove(oFeesLateFeeLocal)
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
        Public Sub Remove(ByVal oLustEvent As MUSTER.Info.FeeLateFeeInfo)
            Try
                oFeesLateFeeCol.Remove(oLustEvent)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LustEvent " & oLustEvent.ID & " is not in the collection of FeeLateFee.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xFeesLateFeeInfo As MUSTER.Info.FeeLateFeeInfo
            For Each xFeesLateFeeInfo In oFeesLateFeeCol.Values
                If xFeesLateFeeInfo.IsDirty Then
                    oFeesLateFee = xFeesLateFeeInfo
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
                    If oFeesLateFee.ID <= 0 Then
                        oFeesLateFee.CreatedBy = UserID
                    Else
                        oFeesLateFee.ModifiedBy = UserID
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
            '    Dim strArr() As String = oFeesLateFeeCol.GetKeys()
            '    Dim nArr(strArr.GetUpperBound(0)) As Integer
            '    Dim y As String
            '    For Each y In strArr
            '        nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            '    Next
            '    nArr.Sort(nArr)
            '    colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            '    If colIndex + direction > -1 And _
            '        colIndex + direction <= nArr.GetUpperBound(0) Then
            '        Return oFeesLateFeeCol.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            '    Else
            '        Return oFeesLateFeeCol.Item(nArr.GetValue(colIndex)).ID.ToString
            '    End If
        End Function
#End Region
#Region "General Operations"

        Public Function Reset() As Object
            oFeesLateFee.Reset()
            oFeesLateFeeCol.Clear()
        End Function
#End Region


#Region " Populate Routines "
        Public Function CheckForExistingCertNumber(ByVal CertNumber As String, ByVal id As Integer, ByVal isRedTag As Boolean, Optional ByVal ownerID As Integer = -1, Optional ByRef ownerSame As Boolean = False) As Boolean
            Dim dtReturn As Boolean
            Dim key As String
            Dim key2 As String
            Dim results As Integer

            If isRedTag Then
                key = String.Format("{0}{1}", "RT", id)
                key2 = String.Format("{0}{1}", "RT", ownerID)
            Else
                key = id.ToString
                key2 = -1
            End If

            Try


                results = oFeesLateFeeDB.DBGetIfCertIsValid(key2, CertNumber, key, results)

                ownerSame = (results < 0)

                If results >= 0 Then
                    Return (results > 0)
                Else
                    Return True
                End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateLateFeeCertificationGrid(Optional ByVal isRedTag As Boolean = False) As DataSet
            Dim dsRemSys As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel1 As DataRelation
            Dim strSQL As String
            Try


                If Not isRedTag Then
                    strSQL = "select * from vFeesLateFeeCertification_Header where INV_AMT <> '0.00' order by Owner_ID;"
                    strSQL = strSQL & " "
                    strSQL = strSQL & "select * from vFeesLateFeeCertification_Detail ;"
                Else
                    strSQL = "select * from vFeesRedTagCertification_Header where [Year] = " + Now.Year.ToString + ";"
                    strSQL = strSQL & " "
                    strSQL = strSQL & "select * from vFeesRedTagCertification_Detail;"

                End If


                dsRemSys = oFeesLateFeeDB.DBGetDS(strSQL)

                If dsRemSys.Tables(1).Rows.Count > 0 AndAlso Not isRedTag Then
                    dsRel1 = New DataRelation("HeaderToLine", dsRemSys.Tables(0).Columns("INV_Number"), dsRemSys.Tables(1).Columns("INV_Number"), False)
                    dsRemSys.Relations.Add(dsRel1)
                ElseIf dsRemSys.Tables(1).Rows.Count > 0 AndAlso isRedTag Then
                    dsRel1 = New DataRelation("HeaderToLine", New System.Data.DataColumn() {dsRemSys.Tables(0).Columns("Owner_ID"), dsRemSys.Tables(0).Columns("Facility_ID")}, New DataColumn() {dsRemSys.Tables(1).Columns("Owner_ID"), dsRemSys.Tables(1).Columns("Facility_ID")}, False)
                    dsRemSys.Relations.Add(dsRel1)

                End If


                Return dsRemSys

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateWaiveLateFeeGrid() As DataSet
            Dim dsRemSys As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel1 As DataRelation
            Dim strSQL As String
            Try


                strSQL = "select * from vFeesWaiveLateFee_Header ;"
                strSQL = strSQL & " "
                strSQL = strSQL & "select * from vFeesWaiveLateFee_Detail ;"

                dsRemSys = oFeesLateFeeDB.DBGetDS(strSQL)

                If dsRemSys.Tables(1).Rows.Count > 0 Then
                    dsRel1 = New DataRelation("HeaderToLine", dsRemSys.Tables(0).Columns("INV_Number"), dsRemSys.Tables(1).Columns("INV_Number"), False)
                    dsRemSys.Relations.Add(dsRel1)
                End If


                Return dsRemSys

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateLateFeeWaiverExcuses() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFeesLateFeeWaiverExcuses", True, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateLateFeeWaiverDecisions() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFeesLateWaiverDecision", True, True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        '


        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal PropertyMaster As Boolean = True, Optional ByVal IncludeBlank As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            strSQL = ""
            If PropertyMaster Then
                If IncludeBlank Then
                    strSQL = " SELECT '' as PROPERTY_NAME, -1 as PROPERTY_ID, -1 as PROPERTY_POSITION "
                    strSQL &= " UNION "
                End If

                strSQL &= " SELECT PROPERTY_NAME, PROPERTY_ID, PROPERTY_POSITION FROM " & strProperty
                strSQL &= " order by 1 "
            Else
                strSQL &= " SELECT * FROM " & strProperty
            End If
            Try
                dsReturn = oFeesLateFeeDB.DBGetDS(strSQL)
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

    End Class
End Namespace
