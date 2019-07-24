' -------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pFinancialCommitAdjustment
' Provides the operations required to manipulate a FinancialCommitAdjustment object.
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0         AB       06/27/2005    Original class definition
' 
' Function          Description
' -------------------------------------------------------------------------------
' Attribute          Description
' -------------------------------------------------------------------------------
Namespace MUSTER.BusinessLogic

    <Serializable()> Public Class pFinancialCommitAdjustment

#Region "Private Member Variables"
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        Private colFinancialCommitAdjustment As MUSTER.Info.FinancialCommitAdjustmentCollection = New MUSTER.Info.FinancialCommitAdjustmentCollection
        Private oFinancialCommitAdjustmentInfo As MUSTER.Info.FinancialCommitAdjustmentInfo = New MUSTER.Info.FinancialCommitAdjustmentInfo
        Private oFinancialCommitAdjustmentDB As MUSTER.DataAccess.FinancialCommitAdjustmentDB = New MUSTER.DataAccess.FinancialCommitAdjustmentDB
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Financial").ID
        Private nID As Int64 = -1
#End Region
#Region "Exposed Events"
        Public Delegate Sub FinancialCommitAdjustmentBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialCommitAdjustmentBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialCommitAdjustmentBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FinancialCommitAdjustmentInfoChanged()
        Public Event FinancialCommitAdjustmentBLChanged As FinancialCommitAdjustmentBLChangedEventHandler
        Public Event FinancialCommitAdjustmentBLColChanged As FinancialCommitAdjustmentBLColChangedEventHandler
        Public Event FinancialCommitAdjustmentBLErr As FinancialCommitAdjustmentBLErrEventHandler
        Public Event FinancialCommitAdjustmentInfChanged As FinancialCommitAdjustmentInfoChanged
        ' indicates change in the underlying _ProtoInfo structure
#End Region
#Region "Constructors"
        Public Sub New()
            oFinancialCommitAdjustmentInfo = New MUSTER.Info.FinancialCommitAdjustmentInfo
        End Sub
        Public Sub New(ByVal TextID As Integer)
            oFinancialCommitAdjustmentInfo = New MUSTER.Info.FinancialCommitAdjustmentInfo
            Me.Retrieve(TextID)
        End Sub

#End Region
#Region "Exposed Attributes"
        ' The unique ID for the row containing the text in the table (info.ID)
        Public ReadOnly Property CommitAdjustmentID() As Int64
            Get
                Return oFinancialCommitAdjustmentInfo.CommitAdjustmentID
            End Get
        End Property


        Public Property CommitmentID() As Int64
            Get
                Return oFinancialCommitAdjustmentInfo.CommitmentID
            End Get
            Set(ByVal Value As Int64)
                oFinancialCommitAdjustmentInfo.CommitmentID = Value
            End Set
        End Property
        Public Property AdjustDate() As Date
            Get
                Return oFinancialCommitAdjustmentInfo.AdjustDate
            End Get
            Set(ByVal Value As Date)
                oFinancialCommitAdjustmentInfo.AdjustDate = Value
            End Set
        End Property
        Public Property AdjustType() As Int64
            Get
                Return oFinancialCommitAdjustmentInfo.AdjustType
            End Get
            Set(ByVal Value As Int64)
                oFinancialCommitAdjustmentInfo.AdjustType = Value
            End Set
        End Property
        Public Property AdjustAmount() As Double
            Get
                Return oFinancialCommitAdjustmentInfo.AdjustMoney
            End Get
            Set(ByVal Value As Double)
                oFinancialCommitAdjustmentInfo.AdjustMoney = Value
            End Set
        End Property
        Public Property DirectorApprovalReq() As Boolean
            Get
                Return oFinancialCommitAdjustmentInfo.DirectorApprovalReq
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitAdjustmentInfo.DirectorApprovalReq = Value
            End Set
        End Property
        Public Property FinancialApprovalReq() As Boolean
            Get
                Return oFinancialCommitAdjustmentInfo.FinancialApprovalReq
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitAdjustmentInfo.FinancialApprovalReq = Value
            End Set
        End Property
        Public Property Approved() As Boolean
            Get
                Return oFinancialCommitAdjustmentInfo.Approved
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitAdjustmentInfo.Approved = Value
            End Set
        End Property
        Public Property Comments() As String
            Get
                Return oFinancialCommitAdjustmentInfo.Comments
            End Get
            Set(ByVal Value As String)
                oFinancialCommitAdjustmentInfo.Comments = Value
            End Set
        End Property

        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oFinancialCommitAdjustmentInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFinancialCommitAdjustmentInfo.CreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFinancialCommitAdjustmentInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oFinancialCommitAdjustmentInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitAdjustmentInfo.Deleted = Value
            End Set
        End Property
        ' The entity ID associated with a financialtext object.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return oFinancialCommitAdjustmentInfo.EntityID
            End Get
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oFinancialCommitAdjustmentInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFinancialCommitAdjustmentInfo.IsDirty = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFinancialCommitAdjustmentInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFinancialCommitAdjustmentInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFinancialCommitAdjustmentInfo.ModifiedOn
            End Get
        End Property
        Public ReadOnly Property ApprovedOriginal() As Boolean
            Get
                Return oFinancialCommitAdjustmentInfo.ApprovedOriginal
            End Get
        End Property


#End Region
#Region "Exposed Operations"
#Region "General Operations"
        Public Sub Clear()
            oFinancialCommitAdjustmentInfo = New MUSTER.Info.FinancialCommitAdjustmentInfo
        End Sub
        Public Sub Reset()
            oFinancialCommitAdjustmentInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        Public Function Retrieve(ByVal CommitAdjustmentID As Int64) As MUSTER.Info.FinancialCommitAdjustmentInfo
            Try
                oFinancialCommitAdjustmentInfo = oFinancialCommitAdjustmentDB.DBGetByID(CommitAdjustmentID)
                If oFinancialCommitAdjustmentInfo.CommitAdjustmentID = 0 Then
                    nID -= 1
                End If

                Return oFinancialCommitAdjustmentInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If Me.ValidateData(strModuleName) Then

                    oFinancialCommitAdjustmentDB.Put(oFinancialCommitAdjustmentInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFinancialCommitAdjustmentInfo.Archive()

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
            Dim xFinancialCommitAdjustmentInfo As MUSTER.Info.FinancialCommitAdjustmentInfo
            Try
                For Each xFinancialCommitAdjustmentInfo In colFinancialCommitAdjustment.Values
                    If xFinancialCommitAdjustmentInfo.IsDirty Then
                        oFinancialCommitAdjustmentInfo = xFinancialCommitAdjustmentInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oFinancialCommitAdjustmentInfoInfo As MUSTER.Info.FinancialCommitAdjustmentInfo)
            Try
                colFinancialCommitAdjustment.Add(oFinancialCommitAdjustmentInfoInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFinancialCommitAdjustmentInfo As Object)
            Try
                colFinancialCommitAdjustment.Remove(oFinancialCommitAdjustmentInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        ' Gets all the info
        Public Function GetAllByCommitment(ByVal CommitmentID As Int64) As MUSTER.Info.FinancialCommitAdjustmentCollection
            Try
                colFinancialCommitAdjustment.Clear()
                colFinancialCommitAdjustment = oFinancialCommitAdjustmentDB.DBGetByCommitment(CommitmentID)
                Return colFinancialCommitAdjustment
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
        '        dsReturn = oFinancialCommitAdjustmentDB.DBGetDS(strSQL)
        '        If dsReturn.Tables(0).Rows.Count > 0 Then
        '            dtReturn = dsReturn.Tables(0)
        '        End If
        '        Return dtReturn
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        'End Function
        Public Function PopulateFinancialCommitmentAdjustmentTypes() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialCommitmentAdjustmentTypes", False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal IncludeBlank As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            strSQL = ""
            If IncludeBlank Then
                strSQL = " SELECT '' as PROPERTY_NAME, 0 as PROPERTY_ID, 0 as PROPERTY_POSITION "
                strSQL &= " UNION "
            End If

            strSQL &= " SELECT PROPERTY_NAME, PROPERTY_ID, PROPERTY_POSITION FROM " & strProperty
            strSQL &= " order by 1 "
            Try
                dsReturn = oFinancialCommitAdjustmentDB.DBGetDS(strSQL)
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
            Dim strArr() As String = colFinancialCommitAdjustment.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.CommitAdjustmentID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colFinancialCommitAdjustment.Item(nArr.GetValue(colIndex + direction)).CommitAdjustmentID.ToString
            Else
                Return colFinancialCommitAdjustment.Item(nArr.GetValue(colIndex)).CommitAdjustmentID.ToString
            End If
        End Function
#End Region

    End Class

End Namespace

