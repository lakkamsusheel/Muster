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
' Function          Description
' -------------------------------------------------------------------------------
' Attribute          Description
' -------------------------------------------------------------------------------
Namespace MUSTER.BusinessLogic

    <Serializable()> Public Class pFinancialInvoice

#Region "Private Member Variables"
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        Private colFinancialInvoice As MUSTER.Info.FinancialInvoiceCollection = New MUSTER.Info.FinancialInvoiceCollection
        Private oFinancialInvoiceInfo As MUSTER.Info.FinancialInvoiceInfo = New MUSTER.Info.FinancialInvoiceInfo
        Private oFinancialInvoiceDB As MUSTER.DataAccess.FinancialInvoiceDB = New MUSTER.DataAccess.FinancialInvoiceDB
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Financial").ID
        Private nID As Int64 = -1
#End Region
#Region "Exposed Events"
        Public Delegate Sub FinancialInvoiceBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialInvoiceBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialInvoiceBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FinancialInvoiceInfoChanged()
        Public Event FinancialInvoiceBLChanged As FinancialInvoiceBLChangedEventHandler
        Public Event FinancialInvoiceBLColChanged As FinancialInvoiceBLColChangedEventHandler
        Public Event FinancialInvoiceBLErr As FinancialInvoiceBLErrEventHandler
        Public Event FinancialInvoiceInfChanged As FinancialInvoiceInfoChanged
        ' indicates change in the underlying _ProtoInfo structure
#End Region
#Region "Constructors"
        Public Sub New()
            oFinancialInvoiceInfo = New MUSTER.Info.FinancialInvoiceInfo
        End Sub
        Public Sub New(ByVal TextID As Integer)
            oFinancialInvoiceInfo = New MUSTER.Info.FinancialInvoiceInfo
            Me.Retrieve(TextID)
        End Sub

#End Region
#Region "Exposed Attributes"

        Public ReadOnly Property id() As Int64
            Get
                Return oFinancialInvoiceInfo.ID
            End Get
        End Property
        Public Property ReimbursementID() As Int64
            Get
                Return oFinancialInvoiceInfo.ReimbursementID
            End Get
            Set(ByVal Value As Int64)
                oFinancialInvoiceInfo.ReimbursementID = Value
            End Set
        End Property
        Public Property PaymentSequence() As Int64
            Get
                Return oFinancialInvoiceInfo.PaymentSequence
            End Get
            Set(ByVal Value As Int64)
                oFinancialInvoiceInfo.PaymentSequence = Value
            End Set
        End Property
        Public Property VendorInvoice() As String
            Get
                Return oFinancialInvoiceInfo.VendorInvoice
            End Get
            Set(ByVal Value As String)
                oFinancialInvoiceInfo.VendorInvoice = Value
            End Set
        End Property
        Public Property InvoicedAmount() As Double
            Get
                Return oFinancialInvoiceInfo.InvoicedAmount
            End Get
            Set(ByVal Value As Double)
                oFinancialInvoiceInfo.InvoicedAmount = Value
            End Set
        End Property
        Public Property PaidAmount() As Double
            Get
                Return oFinancialInvoiceInfo.PaidAmount
            End Get
            Set(ByVal Value As Double)
                oFinancialInvoiceInfo.PaidAmount = Value
            End Set
        End Property

        Public Property DeductionReason() As String
            Get
                Return oFinancialInvoiceInfo.DeductionReason
            End Get
            Set(ByVal Value As String)
                oFinancialInvoiceInfo.DeductionReason = Value
            End Set
        End Property
        Public Property OnHold() As Boolean
            Get
                Return oFinancialInvoiceInfo.OnHold
            End Get
            Set(ByVal Value As Boolean)
                oFinancialInvoiceInfo.OnHold = Value
            End Set
        End Property

        Public Property Final() As Boolean
            Get
                Return oFinancialInvoiceInfo.Final
            End Get
            Set(ByVal Value As Boolean)
                oFinancialInvoiceInfo.Final = Value
            End Set
        End Property
        Public Property Comment() As String
            Get
                Return oFinancialInvoiceInfo.Comment
            End Get
            Set(ByVal Value As String)
                oFinancialInvoiceInfo.Comment = Value
            End Set
        End Property

        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oFinancialInvoiceInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFinancialInvoiceInfo.CreatedBy = Value
            End Set
        End Property

        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFinancialInvoiceInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oFinancialInvoiceInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFinancialInvoiceInfo.Deleted = Value
            End Set
        End Property
        ' The entity ID associated with a financialtext object.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return oFinancialInvoiceInfo.EntityID
            End Get
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oFinancialInvoiceInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFinancialInvoiceInfo.IsDirty = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFinancialInvoiceInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFinancialInvoiceInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFinancialInvoiceInfo.ModifiedOn
            End Get
        End Property

        Public Property PONumber() As String
            Get
                Return oFinancialInvoiceInfo.PONumber
            End Get
            Set(ByVal Value As String)
                oFinancialInvoiceInfo.PONumber = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "General Operations"
        Public Sub Clear()
            oFinancialInvoiceInfo = New MUSTER.Info.FinancialInvoiceInfo
        End Sub
        Public Sub Reset()
            oFinancialInvoiceInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        Public Function Retrieve(ByVal CommitmentID As Int64) As MUSTER.Info.FinancialInvoiceInfo
            Try
                oFinancialInvoiceInfo = oFinancialInvoiceDB.DBGetByID(CommitmentID)
                If oFinancialInvoiceInfo.ID = 0 Then
                    nID -= 1
                End If

                Return oFinancialInvoiceInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If Me.ValidateData(strModuleName) Then

                    oFinancialInvoiceDB.Put(oFinancialInvoiceInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFinancialInvoiceInfo.Archive()

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
        Public Function CommitmentHasOpenChangeOrder(ByVal commitmentID As Int64) As Boolean
            Dim strSQL As String = String.Empty
            Dim retVal As Boolean = False
            Try
                strSQL = "select count(commitmentid) from vFinancialCommitmentAdjustment_Grid " + _
                            "where commitmentid = " + commitmentID.ToString + " " + _
                            "and adjust_type in (" + _
                            "select property_name from vFinancialCommitmentAdjustmentTypes where property_id = 1075) " + _
                            "and (fin_app_req = 1 or director_app_req = 1) " + _
                            "and approved = 0"
                Dim ds As DataSet = oFinancialInvoiceDB.DBGetDS(strSQL)
                If ds.Tables.Count > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        If ds.Tables(0).Rows(0)(0) > 0 Then
                            retVal = True
                        End If
                    End If
                End If
                Return retVal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xFinancialInvoiceInfo As MUSTER.Info.FinancialInvoiceInfo
            Try
                For Each xFinancialInvoiceInfo In colFinancialInvoice.Values
                    If xFinancialInvoiceInfo.IsDirty Then
                        oFinancialInvoiceInfo = xFinancialInvoiceInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oFinancialInvoiceInfoInfo As MUSTER.Info.FinancialInvoiceInfo)
            Try
                colFinancialInvoice.Add(oFinancialInvoiceInfoInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFinancialInvoiceInfo As Object)
            Try
                colFinancialInvoice.Remove(oFinancialInvoiceInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        ' Gets all the info
        Public Function GetAllByReimbursement(ByVal ReimbursementID As Int64) As MUSTER.Info.FinancialInvoiceCollection
            Try
                colFinancialInvoice.Clear()
                colFinancialInvoice = oFinancialInvoiceDB.DBGetByReimbursement(ReimbursementID)
                Return colFinancialInvoice
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
        '        dsReturn = oFinancialInvoiceDB.DBGetDS(strSQL)
        '        If dsReturn.Tables(0).Rows.Count > 0 Then
        '            dtReturn = dsReturn.Tables(0)
        '        End If
        '        Return dtReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
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
                dsReturn = oFinancialInvoiceDB.DBGetDS(strSQL)
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
            Dim strArr() As String = colFinancialInvoice.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.id.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colFinancialInvoice.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colFinancialInvoice.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region

    End Class

End Namespace
