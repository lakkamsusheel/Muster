' -------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pFinancialReimbursement
' Provides the operations required to manipulate a FinancialReimbursement object.
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0         AB       06/24/2005    Original class definition
' 2.0   Thomas Franey  02/25/09     Added Po Number & Comment Functionality into BO
' 
' Function          Description
' -------------------------------------------------------------------------------
' Attribute          Description
' -------------------------------------------------------------------------------
Namespace MUSTER.BusinessLogic

    <Serializable()> Public Class pFinancialReimbursement

#Region "Private Member Variables"
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        Private colFinancialReimbursement As MUSTER.Info.FinancialReimbursementCollection = New MUSTER.Info.FinancialReimbursementCollection
        Private oFinancialReimbursementInfo As MUSTER.Info.FinancialReimbursementInfo = New MUSTER.Info.FinancialReimbursementInfo
        Private oFinancialReimbursementDB As MUSTER.DataAccess.FinancialReimbursementDB = New MUSTER.DataAccess.FinancialReimbursementDB
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Financial").ID
        Private nID As Int64 = -1
#End Region
#Region "Exposed Events"
        Public Delegate Sub FinancialReimbursementBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialReimbursementBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialReimbursementBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FinancialReimbursementInfoChanged()
        Public Event FinancialReimbursementBLChanged As FinancialReimbursementBLChangedEventHandler
        Public Event FinancialReimbursementBLColChanged As FinancialReimbursementBLColChangedEventHandler
        Public Event FinancialReimbursementBLErr As FinancialReimbursementBLErrEventHandler
        Public Event FinancialReimbursementInfChanged As FinancialReimbursementInfoChanged
        ' indicates change in the underlying _ProtoInfo structure
#End Region
#Region "Constructors"
        Public Sub New()
            oFinancialReimbursementInfo = New MUSTER.Info.FinancialReimbursementInfo
        End Sub
        Public Sub New(ByVal TextID As Integer)
            oFinancialReimbursementInfo = New MUSTER.Info.FinancialReimbursementInfo
            Me.Retrieve(TextID)
        End Sub

#End Region
#Region "Exposed Attributes"



        Public ReadOnly Property id() As Int64
            Get
                Return oFinancialReimbursementInfo.ID
            End Get
        End Property
        Public Property FinancialEventID() As Int64
            Get
                Return oFinancialReimbursementInfo.FinancialEventID
            End Get
            Set(ByVal Value As Int64)
                oFinancialReimbursementInfo.FinancialEventID = Value
            End Set
        End Property
        Public Property CommitmentID() As Int64
            Get
                Return oFinancialReimbursementInfo.CommitmentID
            End Get
            Set(ByVal Value As Int64)
                oFinancialReimbursementInfo.CommitmentID = Value
            End Set
        End Property
        Public Property PaymentNumber() As Int64
            Get
                Return oFinancialReimbursementInfo.PaymentNumber
            End Get
            Set(ByVal Value As Int64)
                oFinancialReimbursementInfo.PaymentNumber = Value
            End Set
        End Property
        Public Property ReceivedDate() As Date
            Get
                Return oFinancialReimbursementInfo.ReceivedDate
            End Get
            Set(ByVal Value As Date)
                oFinancialReimbursementInfo.ReceivedDate = Value
            End Set
        End Property
        Public Property PaymentDate() As Date
            Get
                Return oFinancialReimbursementInfo.PaymentDate
            End Get
            Set(ByVal Value As Date)
                oFinancialReimbursementInfo.PaymentDate = Value
            End Set
        End Property
        Public Property RequestedAmount() As Double
            Get
                Return oFinancialReimbursementInfo.RequestedAmount
            End Get
            Set(ByVal Value As Double)
                oFinancialReimbursementInfo.RequestedAmount = Value
            End Set
        End Property
        Public Property IncompleteReason() As String
            Get
                Return oFinancialReimbursementInfo.IncompleteReason
            End Get
            Set(ByVal Value As String)
                oFinancialReimbursementInfo.IncompleteReason = Value
            End Set
        End Property

        Public Property Incomplete() As Boolean
            Get
                Return oFinancialReimbursementInfo.Incomplete
            End Get
            Set(ByVal Value As Boolean)
                oFinancialReimbursementInfo.Incomplete = Value
            End Set
        End Property
        Public Property IncompleteOther() As String
            Get
                Return oFinancialReimbursementInfo.IncompleteOther
            End Get
            Set(ByVal Value As String)
                oFinancialReimbursementInfo.IncompleteOther = Value
            End Set
        End Property
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oFinancialReimbursementInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFinancialReimbursementInfo.CreatedBy = Value
            End Set
        End Property

        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFinancialReimbursementInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oFinancialReimbursementInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFinancialReimbursementInfo.Deleted = Value
            End Set
        End Property
        ' The entity ID associated with a financialtext object.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return oFinancialReimbursementInfo.EntityID
            End Get
        End Property

        Public Property PONumber() As String
            Get
                Return oFinancialReimbursementInfo.PONumber
            End Get
            Set(ByVal Value As String)
                oFinancialReimbursementInfo.PONumber = Value
            End Set
        End Property

        Public Property Comment() As String
            Get
                Return oFinancialReimbursementInfo.Comment
            End Get
            Set(ByVal Value As String)
                oFinancialReimbursementInfo.Comment = Value
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oFinancialReimbursementInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFinancialReimbursementInfo.IsDirty = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFinancialReimbursementInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFinancialReimbursementInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFinancialReimbursementInfo.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "General Operations"
        Public Sub Clear()
            oFinancialReimbursementInfo = New MUSTER.Info.FinancialReimbursementInfo
        End Sub
        Public Sub Reset()
            oFinancialReimbursementInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        Public Function Retrieve(ByVal CommitmentID As Int64) As MUSTER.Info.FinancialReimbursementInfo
            Try
                oFinancialReimbursementInfo = oFinancialReimbursementDB.DBGetByID(CommitmentID)
                If oFinancialReimbursementInfo.ID = 0 Then
                    nID -= 1
                End If

                Return oFinancialReimbursementInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If Me.ValidateData(strModuleName) Then

                    oFinancialReimbursementDB.Put(oFinancialReimbursementInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFinancialReimbursementInfo.Archive()

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
            Dim xFinancialReimbursementInfo As MUSTER.Info.FinancialReimbursementInfo
            Try
                For Each xFinancialReimbursementInfo In colFinancialReimbursement.Values
                    If xFinancialReimbursementInfo.IsDirty Then
                        oFinancialReimbursementInfo = xFinancialReimbursementInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oFinancialReimbursementInfoInfo As MUSTER.Info.FinancialReimbursementInfo)
            Try
                colFinancialReimbursement.Add(oFinancialReimbursementInfoInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFinancialReimbursementInfo As Object)
            Try
                colFinancialReimbursement.Remove(oFinancialReimbursementInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        ' Gets all the info
        Public Function GetAllByFinancialEvent(ByVal CommitmentID As Int64) As MUSTER.Info.FinancialReimbursementCollection
            Try
                colFinancialReimbursement.Clear()
                colFinancialReimbursement = oFinancialReimbursementDB.DBGetByCommitment(CommitmentID)
                Return colFinancialReimbursement
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
        '        dsReturn = oFinancialReimbursementDB.DBGetDS(strSQL)
        '        If dsReturn.Tables(0).Rows.Count > 0 Then
        '            dtReturn = dsReturn.Tables(0)
        '        End If
        '        Return dtReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        'End Function

        Public Function PopulateFinancialIncompleteAppReasons() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialIncompleteAppReasons", False, False)
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
                dsReturn = oFinancialReimbursementDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function ProcessReimbursementNotification(ByVal nVal As String, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Try

                oFinancialReimbursementDB.DBProcessReimbursementNotification(nVal, Now.Date, moduleID, staffID, returnVal, UserID)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetFinancialIncompleteAppReasonsForLetters(ByVal CheckedReasons As String) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "select * from vFinancialIncompleteAppReasons where"
                strSQL += " Text_ID in (" + CheckedReasons.ToString + ")"

                dsReturn = oFinancialReimbursementDB.DBGetDS(strSQL)
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
        Public Function GetNoticeDocuments(ByVal nReimbursmentID As Integer) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "select document_location+document_name as docName from tblsys_document_manager where"
                strSQL += " date_edited in ("
                strSQL += "select max(date_edited) from tblsys_document_manager where entity_id=" + nReimbursmentID.ToString + " and document_description in ('Reimbursement Memo') union "
                strSQL += "select max(date_edited) from tblsys_document_manager where entity_id=" + nReimbursmentID.ToString + " and document_description in ('Notice of Reimbursement')) "
                strSQL += "and entity_id=" + nReimbursmentID.ToString
                dsReturn = oFinancialReimbursementDB.DBGetDS(strSQL)
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
            Dim strArr() As String = colFinancialReimbursement.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.CommitmentID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colFinancialReimbursement.Item(nArr.GetValue(colIndex + direction)).CommitmentID.ToString
            Else
                Return colFinancialReimbursement.Item(nArr.GetValue(colIndex)).CommitmentID.ToString
            End If
        End Function
#End Region

    End Class

End Namespace
