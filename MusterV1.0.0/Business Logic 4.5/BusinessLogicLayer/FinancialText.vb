
Namespace MUSTER.BusinessLogic
    ' ' -------------------------------------------------------------------------------
    ' ' MUSTER.BusinessLogic.pFinancialText
    ' ' Provides the operations required to manipulate a FinancialText object.
    ' ' 
    ' ' Copyright (C) 2004, 2005 CIBER, Inc.
    ' ' All rights reserved.
    ' ' 
    ' ' Release   Initials    Date        Description
    ' ' 1.0         JC       06/08/2005    Original class definition
    ' ' 
    ' ' Function          Description
    ' ' -------------------------------------------------------------------------------
    ' ' Attribute          Description
    ' ' -------------------------------------------------------------------------------
    <Serializable()> Public Class pFinancialText

#Region "Private Member Variables"
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        Private colFinText As MUSTER.Info.FinancialTextCollection = New MUSTER.Info.FinancialTextCollection
        Private oFinTextInfo As MUSTER.Info.FinancialTextInfo = New MUSTER.Info.FinancialTextInfo
        Private oFinTextDB As MUSTER.DataAccess.FinancialTextDB = New MUSTER.DataAccess.FinancialTextDB
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Financial").ID
        Private nID As Int64 = -1
#End Region
#Region "Exposed Events"
        Public Delegate Sub FinTextBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinTextBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinTextBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FinTextInfoChanged()
        Public Event FinTextBLChanged As FinTextBLChangedEventHandler
        Public Event FinTextBLColChanged As FinTextBLColChangedEventHandler
        Public Event FinTextBLErr As FinTextBLErrEventHandler
        Public Event FinTextInfChanged As FinTextInfoChanged
        ' indicates change in the underlying _ProtoInfo structure
#End Region
#Region "Constructors"
        Public Sub New()
            oFinTextInfo = New MUSTER.Info.FinancialTextInfo
        End Sub
        Public Sub New(ByVal TextID As Integer)
            oFinTextInfo = New MUSTER.Info.FinancialTextInfo
            Me.Retrieve(TextID)
        End Sub

#End Region
#Region "Exposed Attributes"
        ' The type associated with the financial text info object (info.reason_type)
        Public Property Text_Type() As Integer
            Get
                Return oFinTextInfo.Reason_Type
            End Get
            Set(ByVal Value As Integer)
                oFinTextInfo.Reason_Type = Value
            End Set
        End Property
        ' The text associated with the instance of the financial text info object (info.reason_text)
        Public Property Text() As String
            Get
                Return oFinTextInfo.Reason_Text
            End Get
            Set(ByVal Value As String)
                oFinTextInfo.Reason_Text = Value
            End Set
        End Property
        ' The "common name" associated with the text (info.reason_name)
        Public Property Name() As String
            Get
                Return oFinTextInfo.Reason_Name
            End Get
            Set(ByVal Value As String)
                oFinTextInfo.Reason_Name = Value
            End Set
        End Property
        ' Gets/Sets the active flag for the financial text info object (from info.Active)
        Public Property Active() As Boolean
            Get
                Return oFinTextInfo.Active
            End Get
            Set(ByVal Value As Boolean)
                oFinTextInfo.Active = Value
            End Set
        End Property

        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oFinTextInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFinTextInfo.CreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFinTextInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oFinTextInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFinTextInfo.Deleted = Value
            End Set
        End Property
        ' The entity ID associated with a financialtext object.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return oFinTextInfo.EntityID
            End Get
        End Property
        ' The unique ID for the row containing the text in the table (info.ID)
        Public ReadOnly Property ID() As Int64
            Get
                Return oFinTextInfo.ID
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oFinTextInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFinTextInfo.IsDirty = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFinTextInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFinTextInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFinTextInfo.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "General Operations"
        Public Sub Clear()
            oFinTextInfo = New MUSTER.Info.FinancialTextInfo
        End Sub
        Public Sub Reset()
            oFinTextInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        Public Function Retrieve(ByVal TextID As Int64) As MUSTER.Info.FinancialTextInfo
            Try
                oFinTextInfo = oFinTextDB.DBGetByID(TextID)
                If oFinTextInfo.ID = 0 Then
                    nID -= 1
                End If

                Return oFinTextInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If Me.ValidateData(strModuleName) Then

                    oFinTextDB.Put(oFinTextInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFinTextInfo.Archive()

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
            Dim xFinancialTextInfo As MUSTER.Info.FinancialTextInfo
            Try
                For Each xFinancialTextInfo In colFinText.Values
                    If xFinancialTextInfo.IsDirty Then
                        oFinTextInfo = xFinancialTextInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oFinTextInfoInfo As MUSTER.Info.FinancialTextInfo)
            Try
                colFinText.Add(oFinTextInfoInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFinTextInfo As Object)
            Try
                colFinText.Remove(oFinTextInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        ' Gets all the info
        Public Function GetAllByReasonType(ByVal nReason As Int64) As MUSTER.Info.FinancialTextCollection
            Try
                colFinText.Clear()
                colFinText = oFinTextDB.DBGetByReasonType(nReason)
                Return colFinText
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function


#End Region
#Region " Populate Routines "


        Public Function GetFinancialTextTable(ByVal TextType As Integer) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String


            strSQL = "select * from tblSYS_Text where Reason_Type = " & TextType
            strSQL &= " and deleted = 0 Order By Text_Name"

            Try
                dsReturn = oFinTextDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
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
                dsReturn = oFinTextDB.DBGetDS(strSQL)
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
            Dim strArr() As String = colFinText.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colFinText.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colFinText.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region

    End Class
End Namespace
