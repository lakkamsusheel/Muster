'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.LicenseeComplianceEvent
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'
' Function          Description
' GetEntity(NAME)   Returns the Entity requested by the string arg NAME
' GetEntity(ID)     Returns the Entity requested by the int arg ID
' GetAll()          Returns an ReportsCollection with all Entity objects
' Add(ID)           Adds the Entity identified by arg ID to the 
'                           internal ReportsCollection
' Add(Name)         Adds the Entity identified by arg NAME to the internal 
'                           ReportsCollection
' Add(Entity)       Adds the Entity passed as the argument to the internal 
'                           ReportsCollection
' Remove(ID)        Removes the Entity identified by arg ID from the internal 
'                           ReportsCollection
' Remove(NAME)      Removes the Entity identified by arg NAME from the 
'                           internal ReportsCollection
' EntityTable()     Returns a datatable containing all columns for the Entity 
'                           objects in the internal ReportsCollection.
'
' NOTE: This file to be used as LicenseeComplianceEvent to build other objects.
'       Replace keyword "FacilityComplianceEvent" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pLicenseeComplianceEvent
#Region "Public Events"
        Public Event LicenseeComplianceEventErr(ByVal MsgStr As String)
        Public Event evtLCEInfoChanged(ByVal bolValue As Boolean)
        Public Event evtLCECollectionChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oLCEInfo As MUSTER.Info.LicenseeComplianceEventInfo
        Private WithEvents colLCE As MUSTER.Info.LicenseeComplianceEventCollection
        Private oLCEDB As New MUSTER.DataAccess.LicenseeComplianceEventDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private strLicenseeName As String = String.Empty
        Private nCitationID As Integer = 0
        Private WithEvents oProperty As New MUSTER.BusinessLogic.pProperty
#End Region
#Region "Constructors"
        Public Sub New()
            oLCEInfo = New MUSTER.Info.LicenseeComplianceEventInfo
            colLCE = New MUSTER.Info.LicenseeComplianceEventCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named FacilityComplianceEvent object.
        '
        '********************************************************
        Public Sub New(ByVal LCEName As String)
            oLCEInfo = New MUSTER.Info.LicenseeComplianceEventInfo
            colLCE = New MUSTER.Info.LicenseeComplianceEventCollection
            Me.Retrieve(LCEName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ColLCEvents() As MUSTER.Info.LicenseeComplianceEventCollection
            Get
                Return colLCE
            End Get
            Set(ByVal Value As MUSTER.Info.LicenseeComplianceEventCollection)
                colLCE = Value
            End Set
        End Property
        Public Property EscalationName() As String
            Get
                Return oLCEInfo.EscalationName
            End Get
            Set(ByVal Value As String)
                oLCEInfo.EscalationName = Value
            End Set
        End Property
        Public Property citationText() As String
            Get
                Return oLCEInfo.CitationText
            End Get
            Set(ByVal Value As String)
                oLCEInfo.CitationText = Value
            End Set
        End Property
        Public Property LCEInfo() As MUSTER.Info.LicenseeComplianceEventInfo
            Get
                Return oLCEInfo
            End Get
            Set(ByVal Value As MUSTER.Info.LicenseeComplianceEventInfo)
                oLCEInfo = Value
            End Set
        End Property
        Public Property LCECollection() As MUSTER.Info.LicenseeComplianceEventCollection
            Get
                Return colLCE
            End Get
            Set(ByVal Value As MUSTER.Info.LicenseeComplianceEventCollection)
                colLCE = Value
            End Set
        End Property
        Public Property LicenseeName() As String
            Get
                Return oLCEInfo.LicenseeName
            End Get
            Set(ByVal Value As String)
                oLCEInfo.LicenseeName = Value
            End Set
        End Property
        Public Property FacilityName() As String
            Get
                Return oLCEInfo.facilityName
            End Get
            Set(ByVal Value As String)
                oLCEInfo.facilityName = Value
            End Set
        End Property
        Public Property ID() As Int32
            Get
                Return oLCEInfo.ID
            End Get
            Set(ByVal Value As Int32)
                oLCEInfo.ID = Value
            End Set
        End Property
        'Public Property citationID() As Integer
        '    Get
        '        Return nCitationID
        '    End Get
        '    Set(ByVal Value As Integer)
        '        nCitationID = Value
        '    End Set
        'End Property
        Public Property LicenseeID() As Int32
            Get
                Return oLCEInfo.LicenseeID
            End Get
            Set(ByVal Value As Int32)
                oLCEInfo.LicenseeID = Value
            End Set
        End Property
        Public Property FacilityID() As Int32
            Get
                Return oLCEInfo.FacilityID
            End Get
            Set(ByVal Value As Int32)
                oLCEInfo.FacilityID = Value
            End Set
        End Property
        Public Property LicenseeCitationID() As Int32
            Get
                Return oLCEInfo.LicenseeCitationID
            End Get
            Set(ByVal Value As Int32)
                oLCEInfo.LicenseeCitationID = Value
            End Set
        End Property
        Public Property CitationDueDate() As Date
            Get
                Return oLCEInfo.CitationDueDate
            End Get

            Set(ByVal Value As Date)
                oLCEInfo.CitationDueDate = Value
            End Set
        End Property
        Public Property CitationReceivedDate() As Date
            Get
                Return oLCEInfo.CitationReceivedDate
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.CitationReceivedDate = Value
            End Set
        End Property
        Public Property Rescinded() As Boolean
            Get
                Return oLCEInfo.Rescinded
            End Get
            Set(ByVal Value As Boolean)
                oLCEInfo.Rescinded = Value
            End Set
        End Property
        Public Property LCEDate() As Date
            Get
                Return oLCEInfo.LCEDate
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.LCEDate = Value
            End Set
        End Property
        Public Property LCEProcessDate() As Date
            Get
                Return oLCEInfo.LCEProcessDate
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.LCEProcessDate = Value
            End Set
        End Property
        Public Property NextDueDate() As Date
            Get
                Return oLCEInfo.NextDueDate
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.NextDueDate = Value
            End Set
        End Property
        Public Property OverrideDueDate() As Date
            Get
                Return oLCEInfo.OverrideDueDate
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.OverrideDueDate = Value
            End Set
        End Property
        Public Property LCEStatus() As Integer
            Get
                Return oLCEInfo.LCEStatus
            End Get
            Set(ByVal Value As Integer)
                oLCEInfo.LCEStatus = Value
            End Set
        End Property
        Public Property Status() As String
            Get
                Return oLCEInfo.Status
            End Get
            Set(ByVal Value As String)
                oLCEInfo.Status = Value
            End Set
        End Property
        Public Property Escalation() As Integer
            Get
                Return oLCEInfo.Escalation
            End Get
            Set(ByVal Value As Integer)
                oLCEInfo.Escalation = Value
            End Set
        End Property
        Public Property PolicyAmount() As Decimal
            Get
                Return oLCEInfo.PolicyAmount
            End Get
            Set(ByVal Value As Decimal)
                oLCEInfo.PolicyAmount = Value
            End Set
        End Property
        Public Property OverrideAmount() As Decimal
            Get
                Return oLCEInfo.OverrideAmount
            End Get
            Set(ByVal Value As Decimal)
                oLCEInfo.OverrideAmount = Value
            End Set
        End Property
        Public Property SettlementAmount() As Decimal
            Get
                Return oLCEInfo.SettlementAmount
            End Get
            Set(ByVal Value As Decimal)
                oLCEInfo.SettlementAmount = Value
            End Set
        End Property
        Public Property PaidAmount() As Decimal
            Get
                Return oLCEInfo.PaidAmount
            End Get
            Set(ByVal Value As Decimal)
                oLCEInfo.PaidAmount = Value
            End Set
        End Property
        Public Property DateReceived() As Date
            Get
                Return oLCEInfo.DateReceived
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.DateReceived = Value
            End Set
        End Property
        Public Property WorkShopDate() As Date
            Get
                Return oLCEInfo.WorkShopDate
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.WorkShopDate = Value
            End Set
        End Property
        Public Property WorkshopResult() As Integer
            Get
                Return oLCEInfo.WorkshopResult
            End Get
            Set(ByVal Value As Integer)
                oLCEInfo.WorkshopResult = Value
            End Set
        End Property
        Public Property ShowCauseDate() As Date
            Get
                Return oLCEInfo.ShowCauseDate
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.ShowCauseDate = Value
            End Set
        End Property
        Public Property ShowCauseResults() As Integer
            Get
                Return oLCEInfo.ShowCauseResults
            End Get
            Set(ByVal Value As Integer)
                oLCEInfo.ShowCauseResults = Value
            End Set
        End Property
        Public Property CommissionDate() As Date
            Get
                Return oLCEInfo.CommissionDate
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.CommissionDate = Value
            End Set
        End Property
        Public Property CommissionResults() As Integer
            Get
                Return oLCEInfo.CommissionResults
            End Get
            Set(ByVal Value As Integer)
                oLCEInfo.CommissionResults = Value
            End Set
        End Property
        Public Property PendingLetter() As Integer
            Get
                Return oLCEInfo.PendingLetter
            End Get
            Set(ByVal Value As Integer)
                oLCEInfo.PendingLetter = Value
            End Set
        End Property
        Public Property LetterGenerated() As Date
            Get
                Return oLCEInfo.LetterGenerated
            End Get
            Set(ByVal Value As Date)
                oLCEInfo.LetterGenerated = Value
            End Set
        End Property
        Public Property LetterPrinted() As Boolean
            Get
                Return oLCEInfo.LetterPrinted
            End Get
            Set(ByVal Value As Boolean)
                oLCEInfo.LetterPrinted = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oLCEInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oLCEInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oLCEInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oLCEInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xLCEinfo As MUSTER.Info.LicenseeComplianceEventInfo
                For Each xLCEinfo In colLCE.Values
                    If xLCEinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oLCEInfo.IsDirty = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oLCEInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oLCEInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oLCEInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oLCEInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oLCEInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oLCEInfo.ModifiedOn
            End Get
        End Property
        Public Property PendingLetterName() As String
            Get
                Return oLCEInfo.PendingLetterName
            End Get
            Set(ByVal Value As String)
                oLCEInfo.PendingLetterName = Value
            End Set
        End Property
        Public Property ShowCauseResultName() As String
            Get
                Return oLCEInfo.ShowCauseResultName
            End Get
            Set(ByVal Value As String)
                oLCEInfo.ShowCauseResultName = Value
            End Set
        End Property
        Public Property WorkshopResultName() As String
            Get
                Return oLCEInfo.WorkshopResultName
            End Get
            Set(ByVal Value As String)
                oLCEInfo.WorkshopResultName = Value
            End Set
        End Property
        Public Property OwnerID() As Integer
            Get
                Return oLCEInfo.OwnerID
            End Get
            Set(ByVal Value As Integer)
                oLCEInfo.OwnerID = Value
            End Set
        End Property
        Public Property OwnerName() As String
            Get
                Return oLCEInfo.OwnerName
            End Get
            Set(ByVal Value As String)
                oLCEInfo.OwnerName = Value
            End Set
        End Property
        Public Property PendingLetterTemplateNum() As Integer
            Get
                Return oLCEInfo.PendingLetterTemplateNum
            End Get
            Set(ByVal Value As Integer)
                oLCEInfo.PendingLetterTemplateNum = Value
            End Set
        End Property

        Public Property CommissionResultsName() As String
            Get
                Return oLCEInfo.CommissionResultName
            End Get
            Set(ByVal Value As String)
                oLCEInfo.CommissionResultName = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.LicenseeComplianceEventInfo
            Dim oLCEInfoLocal As MUSTER.Info.LicenseeComplianceEventInfo
            Try
                For Each oLCEInfoLocal In colLCE.Values
                    If oLCEInfoLocal.ID = ID Then
                        oLCEInfo = oLCEInfoLocal
                        Return oLCEInfo
                    End If
                Next
                oLCEInfo = oLCEDB.DBGetByID(ID)
                colLCE.Add(oLCEInfo)
                Return oLCEInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Public Function Retrieve(ByVal LCEName As String) As MUSTER.Info.LicenseeComplianceEventInfo
        '    Try
        '        oLCEInfo = Nothing
        '        If colLCE.Contains(LCEName) Then
        '            oLCEInfo = colLCE(LCEName)
        '        Else
        '            If oLCEInfo Is Nothing Then
        '                oLCEInfo = New MUSTER.Info.LicenseeComplianceEventInfo
        '            End If
        '            oLCEInfo = oLCEDB.DBGetByName(LCEName)
        '            colLCE.Add(oLCEInfo)
        '        End If
        '        Return oLCEInfo
        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try

        'End Function

        'Private Sub CheckLicenseeRevoke()
        '    Dim oLCEInfoLocal As MUSTER.Info.LicenseeComplianceEventInfo
        '    For Each oLCEInfoLocal In colLCE.Values
        '        If oLCEInfoLocal.LicenseeID = oLCEInfo.LicenseeID And (DateDiff(DateInterval.Year, oLCEInfoLocal.LCEDate, Now.Today) < 3) And Not oLCEInfoLocal.Rescinded And Not oLCEInfoLocal.Deleted Then
        '            Dim oLicensee As New MUSTER.BusinessLogic.pLicensee
        '            oLicensee.Retrieve(oLCEInfoLocal.LicenseeID)
        '            oLicensee.STATUS_ID = "REVOKED"
        '            oLicensee.Save(True)
        '            MsgBox("Licensee " + oLCEInfo.LicenseeID.ToString + " is Revoked")
        '            Exit Sub
        '        End If
        '    Next
        'End Sub


        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False)

            Try
                Dim OldKey As String = oLCEInfo.ID.ToString
                oLCEDB.Put(oLCEInfo, moduleID, staffID, returnVal)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                If oLCEInfo.ID.ToString <> OldKey Then
                    colLCE.ChangeKey(OldKey, oLCEInfo.ID.ToString)
                End If
                oLCEInfo.Archive()
                oLCEInfo.IsDirty = False
                'End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData() As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = False

            Try
                'If oLCEInfo.ID <> 0 Then
                '    If oLCEInfo.LicenseeID <> 0 Then
                '        If oLCEInfo.FacilityID <> 0 Then
                '            If Date.Compare(oLCEInfo.CitationDueDate, CDate("01/01/0001")) = 0 Then
                '                errStr += "Citation Due Date cannot be empty" + vbCrLf
                '                validateSuccess = False
                '            Else
                '                validateSuccess = True
                '            End If
                '        Else
                '            errStr += "FacilityID cannot be empty" + vbCrLf
                '            validateSuccess = False
                '        End If
                '    Else
                '        errStr += "LicenseeID cannot be empty" + vbCrLf
                '        validateSuccess = False
                '    End If
                'End If

                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent LicenseeComplianceEventErr(errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        'Gets all the info
        Public Function GetAll() As MUSTER.Info.LicenseeComplianceEventCollection
            Try
                colLCE.Clear()
                colLCE = oLCEDB.GetAllInfo
                Return colLCE
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oLCEInfo = oLCEDB.DBGetByID(ID)
                If oLCEInfo.ID = 0 Then
                    oLCEInfo.ID = nID
                    nID -= 1
                End If
                colLCE.Add(oLCEInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oLCE As MUSTER.Info.LicenseeComplianceEventInfo)
            Try
                oLCEInfo = oLCE
                colLCE.Add(oLCEInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oLCEInfoLocal As MUSTER.Info.LicenseeComplianceEventInfo

            Try
                For Each oLCEInfoLocal In colLCE.Values
                    If oLCEInfoLocal.ID = ID Then
                        colLCE.Remove(oLCEInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Licensee Compliance Event " & ID.ToString & " is not in the collection of LicenseeComplianceEvents.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oLCE As MUSTER.Info.LicenseeComplianceEventInfo)
            Try
                colLCE.Remove(oLCE)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Licensee Compliance Event " & oLCE.ID & " is not in the collection of LicenseeComplianceEvents.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xLCEInfo As MUSTER.Info.LicenseeComplianceEventInfo
            For Each xLCEInfo In colLCE.Values
                If xLCEInfo.IsDirty Then
                    oLCEInfo = xLCEInfo
                    Me.Save(moduleID, staffID, returnVal)
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
            Dim strArr() As String = colLCE.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colLCE.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colLCE.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oLCEInfo = New MUSTER.Info.LicenseeComplianceEventInfo
        End Sub
        Public Sub Reset()
            oLCEInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function GetEnforcementHistory(ByVal nLCEId As Integer) As MUSTER.Info.LicenseeComplianceEventCollection
            Try
                colLCE.Clear()
                colLCE = oLCEDB.GetLCEEnforcementHistory(nLCEId)
                Return colLCE
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        ' Used for Getting Escalation
        Public Function GetEscalation(ByVal lceInfo As MUSTER.Info.LicenseeComplianceEventInfo) As String
            Dim retVal As String
            Dim dtDate As Date
            Try
                If lceInfo.PendingLetter = 0 Then
                    If Date.Compare(lceInfo.OverrideDueDate, CDate("01/01/0001")) = 0 Then
                        dtDate = lceInfo.NextDueDate
                    Else
                        dtDate = lceInfo.OverrideDueDate
                    End If
                    If Date.Compare(dtDate, Now.Date) > 0 Then
                        retVal = lceInfo.Status

                        'ElseIf Date.Compare(lceInfo.NextDueDate, Now.Date) > 0 Then

                        '    lceInfo.Escalation = lceInfo.LCEStatus '995
                    Else
                        If (lceInfo.Status = "Show Cause Hearing".ToUpper And (lceInfo.ShowCauseDate = CDate("01/01/0001") Or Date.Compare(lceInfo.ShowCauseDate, Now.Date) > 0 Or lceInfo.ShowCauseResults = 0)) Or _
                                                (lceInfo.Status = "Commission hearing".ToUpper And (lceInfo.CommissionDate = CDate("01/01/0001") Or Date.Compare(lceInfo.CommissionDate, Now.Date) > 0 Or lceInfo.CommissionResults = 0)) Or _
                                                (lceInfo.Status = "NFA Rescind".ToUpper) Or _
                                                (lceInfo.Status = "Legal".ToUpper) Or _
                                                (lceInfo.CitationDueDate = CDate("01/01/0001") Or Date.Compare(lceInfo.CitationDueDate, Now.Date) > 0) Then
                            retVal = lceInfo.Status
                        Else
                            'go to NOV
                            retVal = GetNOVLogic(lceInfo)
                        End If
                    End If
                Else
                    retVal = lceInfo.Status
                End If
                Return retVal
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetAmountDue(ByVal lceInfo As MUSTER.Info.LicenseeComplianceEventInfo) As Decimal
            Dim nAmountDue As Decimal = -1.0
            Try
                If lceInfo.OverrideAmount = -1.0 Then
                    If lceInfo.SettlementAmount = -1.0 Then
                        If lceInfo.PolicyAmount <> -1.0 Then
                            nAmountDue = lceInfo.PolicyAmount
                        End If
                    Else
                        nAmountDue = lceInfo.SettlementAmount
                    End If
                Else
                    nAmountDue = lceInfo.OverrideAmount
                End If
                Return nAmountDue
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetNOVLogic(ByVal lceInfo As MUSTER.Info.LicenseeComplianceEventInfo) As String
            Dim nAmountDue As Decimal
            Try
                If lceInfo.Status.ToUpper = "NEW" Then
                    ' check if all citations received and penalty paid
                    nAmountDue = GetAmountDue(lceInfo)
                    If (lceInfo.Rescinded = False And Date.Compare(lceInfo.CitationReceivedDate, CDate("01/01/0001")) <> 0) And nAmountDue <= lceInfo.PaidAmount Then 'And (lceInfo.OverrideAmount = 0 And lceInfo.SettlementAmount = 0 And lceInfo.PolicyAmount = 0) Then
                        '(Not lceInfo.PaidAmount = 0) Then
                        ' GO TO NFA (sheet 4)
                        Return GetNFALogic(lceInfo)
                    Else
                        Return "Show Cause Hearing"
                    End If
                Else
                    ' Go to Hearings
                    Return GetHearingsLogic(lceInfo)
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetHearingsLogic(ByVal lceInfo As MUSTER.Info.LicenseeComplianceEventInfo) As String
            Dim nAmountDue As Decimal
            Try
                nAmountDue = GetAmountDue(lceInfo)
                If lceInfo.LCEStatus = 996 Then
                    If lceInfo.ShowCauseResults = 1003 Then
                        Return "Commission Hearing"
                    Else
                        Return "Show Cause Agreed Order"
                    End If
                ElseIf lceInfo.LCEStatus = 1118 Then
                    ' check if all citations received and penalty paid
                    If (lceInfo.Rescinded = False And Date.Compare(lceInfo.CitationReceivedDate, CDate("01/01/0001")) <> 0) And nAmountDue <= lceInfo.PaidAmount Then '(lceInfo.OverrideAmount = 0 And lceInfo.SettlementAmount = 0 And lceInfo.PolicyAmount = 0) Then '(Not lceInfo.PaidAmount = 0) Then
                        ' GO TO NFA (sheet 4)
                        Return GetNFALogic(lceInfo)
                    Else
                        Return "Commission Hearing"
                    End If
                ElseIf lceInfo.LCEStatus = 997 Then
                    If lceInfo.CommissionResults = 1004 Then
                        Return "Administrative Order"
                    Else
                        Return "NFA"
                    End If
                ElseIf lceInfo.LCEStatus = 999 Then
                    ' check if all citations received and penalty paid
                    If (lceInfo.Rescinded = False And Date.Compare(lceInfo.CitationReceivedDate, CDate("01/01/0001")) <> 0) And nAmountDue <= lceInfo.PaidAmount Then '(lceInfo.OverrideAmount = 0 And lceInfo.SettlementAmount = 0 And lceInfo.PolicyAmount = 0) Then '(Not lceInfo.PaidAmount = 0) Then
                        ' GO TO NFA (sheet 4)
                        Return GetNFALogic(lceInfo)
                    Else
                        Return "Legal"
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetNFALogic(ByVal lceInfo As MUSTER.Info.LicenseeComplianceEventInfo) As String
            Try
                Return "NFA"
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        ' Used for Processing Escalation
        Public Function EscalationLogic(ByRef LCEInfo As MUSTER.Info.LicenseeComplianceEventInfo) As String
            Dim str As String = String.Empty
            Dim dtDate As Date
            Try
                oLCEInfo = LCEInfo
                If oLCEInfo.PendingLetter = 0 Then
                    If Date.Compare(oLCEInfo.OverrideDueDate, CDate("01/01/0001")) = 0 Then
                        dtDate = oLCEInfo.NextDueDate
                    Else
                        dtDate = oLCEInfo.OverrideDueDate
                    End If
                    If Date.Compare(dtDate, Now.Date) > 0 Then
                        oLCEInfo.Escalation = oLCEInfo.LCEStatus

                        'ElseIf Date.Compare(oLCEInfo.NextDueDate, Now.Date) > 0 Then

                        '    oLCEInfo.Escalation = oLCEInfo.LCEStatus '995
                    Else
                        If (oLCEInfo.Status = "Show Cause Hearing".ToUpper And (oLCEInfo.ShowCauseDate = CDate("01/01/0001") Or Date.Compare(oLCEInfo.ShowCauseDate, Now.Date) > 0 Or oLCEInfo.ShowCauseResults = 0)) Or _
                                                (oLCEInfo.Status = "Commission hearing".ToUpper And (oLCEInfo.CommissionDate = CDate("01/01/0001") Or Date.Compare(oLCEInfo.CommissionDate, Now.Date) > 0 Or oLCEInfo.CommissionResults = 0)) Or _
                                                (oLCEInfo.Status = "NFA Rescind".ToUpper) Or _
                                                (oLCEInfo.Status = "Legal".ToUpper) Or _
                                                (oLCEInfo.CitationDueDate = CDate("01/01/0001") Or Date.Compare(oLCEInfo.CitationDueDate, Now.Date) > 0) Then

                            oLCEInfo.Escalation = oLCEInfo.LCEStatus
                        Else
                            'go to NOV
                            str = NOVLogic()
                        End If
                    End If
                Else
                    oLCEInfo.Escalation = oLCEInfo.LCEStatus
                End If
                Return str
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function AmountDue() As Decimal
            Dim nAmountDue As Decimal = -1.0
            Try
                If oLCEInfo.OverrideAmount = -1.0 Then
                    If oLCEInfo.SettlementAmount = -1.0 Then
                        If oLCEInfo.PolicyAmount <> -1.0 Then
                            nAmountDue = oLCEInfo.PolicyAmount
                        End If
                    Else
                        nAmountDue = oLCEInfo.SettlementAmount
                    End If
                Else
                    nAmountDue = oLCEInfo.OverrideAmount
                End If
                Return nAmountDue
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function NOVLogic() As String
            Dim nAmountDue As Decimal
            Try
                If oLCEInfo.Status.ToUpper = "NEW" Then
                    ' check if all citations received and penalty paid
                    nAmountDue = AmountDue()
                    If (oLCEInfo.Rescinded = False And Date.Compare(oLCEInfo.CitationReceivedDate, CDate("01/01/0001")) <> 0) And nAmountDue <= oLCEInfo.PaidAmount Then 'And (oLCEInfo.OverrideAmount = 0 And oLCEInfo.SettlementAmount = 0 And oLCEInfo.PolicyAmount = 0) Then
                        '(Not oLCEInfo.PaidAmount = 0) Then
                        ' GO TO NFA (sheet 4)
                        NFALogic()
                    Else
                        oLCEInfo.LCEStatus = 996
                        oLCEInfo.Status = "show cause hearing".ToUpper
                        oLCEInfo.Escalation = oLCEInfo.LCEStatus
                        Return "Show Cause Hearing|1294" ' 1294 = pending letter template num property id
                        'lceinfo.ShowCauseDate = prompt user (cancel escalation if null)
                        'If oLCEInfo.ShowCauseDate = CDate("01/01/0001") Then
                        '    MsgBox("Show Cause Hearing Date cannot be null. Escalation process is cancelled")
                        '    LCEInfo.Reset()
                        '    Exit Function
                        'End If
                        'oLCEInfo.NextDueDate = oLCEInfo.ShowCauseDate
                        'oLCEInfo.PendingLetter = 1006
                        'oLCEInfo.PendingLetterName = "show cause hearing".ToUpper
                    End If
                Else
                    ' Go to Hearings
                    Return HearingsLogic()
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function HearingsLogic() As String
            Dim nAmountDue As Decimal
            Try
                nAmountDue = AmountDue()
                If oLCEInfo.LCEStatus = 996 Then
                    If oLCEInfo.ShowCauseResults = 1003 Then
                        oLCEInfo.Status = "Commission Hearing".ToUpper
                        oLCEInfo.LCEStatus = 997
                        oLCEInfo.Escalation = oLCEInfo.LCEStatus
                        'Commission hearing date = prompt user (cancel escalation if null)
                        Return "Commission Hearing|1295"
                        'oLCEInfo.NextDueDate = oLCEInfo.CommissionDate
                        'oLCEInfo.PendingLetter = 1008
                        'oLCEInfo.PendingLetterName = "Commission hearing".ToUpper
                    Else
                        oLCEInfo.LCEStatus = 1118
                        oLCEInfo.Escalation = oLCEInfo.LCEStatus
                        oLCEInfo.Status = "Show cause agreed order".ToUpper
                        oLCEInfo.NextDueDate = Now.Today.AddDays(30)
                        oLCEInfo.PendingLetter = 1007
                        oLCEInfo.PendingLetterTemplateNum = 1296
                        oLCEInfo.PendingLetterName = "Show Cause agreed order".ToUpper
                        oLCEInfo.OverrideDueDate = CDate("01/01/0001")
                    End If
                ElseIf oLCEInfo.LCEStatus = 1118 Then
                    ' check if all citations received and penalty paid
                    If (oLCEInfo.Rescinded = False And Date.Compare(oLCEInfo.CitationReceivedDate, CDate("01/01/0001")) <> 0) And nAmountDue <= oLCEInfo.PaidAmount Then '(oLCEInfo.OverrideAmount = 0 And oLCEInfo.SettlementAmount = 0 And oLCEInfo.PolicyAmount = 0) Then '(Not oLCEInfo.PaidAmount = 0) Then
                        ' GO TO NFA (sheet 4)
                        NFALogic()
                    Else
                        oLCEInfo.LCEStatus = 997
                        oLCEInfo.Escalation = oLCEInfo.LCEStatus
                        oLCEInfo.Status = "Commission Hearing".ToUpper
                        'Commission hearing date = prompt user (cancel escalation if null)
                        Return "Commission Hearing|1297"
                        'oLCEInfo.NextDueDate = oLCEInfo.CommissionDate
                        'oLCEInfo.PendingLetter = 1008
                        'oLCEInfo.PendingLetterName = "Commission hearing".ToUpper
                    End If
                ElseIf oLCEInfo.LCEStatus = 997 Then
                    If oLCEInfo.CommissionResults = 1004 Then
                        oLCEInfo.CommissionResults = 1004
                        oLCEInfo.CommissionResultName = "Administrtaive order".ToUpper
                        oLCEInfo.NextDueDate = Now.Today.AddDays(30)
                        oLCEInfo.PendingLetter = 1010
                        oLCEInfo.PendingLetterTemplateNum = 1299
                        oLCEInfo.PendingLetterName = "Admistrative Order".ToUpper
                        oLCEInfo.OverrideDueDate = CDate("01/01/0001")
                        oLCEInfo.LCEStatus = 999
                        oLCEInfo.Escalation = oLCEInfo.LCEStatus
                        oLCEInfo.Status = "Admistrative Order".ToUpper
                    Else
                        oLCEInfo.LCEStatus = 1000
                        oLCEInfo.Escalation = oLCEInfo.LCEStatus
                        oLCEInfo.Status = "NFA".ToUpper
                        oLCEInfo.NextDueDate = Now.AddDays(30)
                        oLCEInfo.PendingLetter = 1009
                        oLCEInfo.PendingLetterTemplateNum = 1298
                        oLCEInfo.PendingLetterName = "Commission Hearing NFA rescinded".ToUpper
                        oLCEInfo.OverrideDueDate = CDate("01/01/0001")
                    End If
                ElseIf oLCEInfo.LCEStatus = 999 Then
                    ' check if all citations received and penalty paid
                    If (oLCEInfo.Rescinded = False And Date.Compare(oLCEInfo.CitationReceivedDate, CDate("01/01/0001")) <> 0) And nAmountDue <= oLCEInfo.PaidAmount Then '(oLCEInfo.OverrideAmount = 0 And oLCEInfo.SettlementAmount = 0 And oLCEInfo.PolicyAmount = 0) Then '(Not oLCEInfo.PaidAmount = 0) Then
                        ' GO TO NFA (sheet 4)
                        NFALogic()
                    Else
                        oLCEInfo.LCEStatus = 998
                        oLCEInfo.Escalation = oLCEInfo.LCEStatus
                        oLCEInfo.Status = "LEGAL".ToUpper
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub NFALogic()
            Try
                oLCEInfo.LCEStatus = 1000
                oLCEInfo.Escalation = oLCEInfo.LCEStatus
                oLCEInfo.Status = "NFA".ToUpper
                oLCEInfo.PendingLetter = 1119
                oLCEInfo.PendingLetterTemplateNum = 1300
                oLCEInfo.PendingLetterName = "NFA Letter".ToUpper
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function EntityTable(Optional ByVal bolLCE As Boolean = False, Optional ByVal bolEnforcementhistory As Boolean = False) As DataSet
            Dim oLCEInfoLocal As MUSTER.Info.LicenseeComplianceEventInfo
            Dim dr, dr1, dr2 As DataRow
            Dim dt, dt1, dt2 As DataTable
            Dim ds As New DataSet
            Dim dsRel1, dsrel2 As DataRelation
            Dim bolLCEExists As Boolean = False
            Dim bolEnfHistoryExists As Boolean = False
            Try
                If colLCE.Count = 0 Then
                    Return Nothing
                End If
                dt = New DataTable
                dt.Columns.Add("LCE_Status", GetType(String))

                dt1 = New DataTable
                dt1.Columns.Add("LCE_ID", GetType(Integer))
                dt1.Columns.Add("Licensee", GetType(String))
                dt1.Columns.Add("Licensee_id", GetType(Integer))
                dt1.Columns.Add("Rescinded", GetType(Boolean))
                dt1.Columns.Add("Facility_ID", GetType(Integer))
                dt1.Columns.Add("LCE_DATE", GetType(Date))
                dt1.Columns.Add("LCE_PROCESS_DATE", GetType(Date))
                dt1.Columns.Add("NEXT_DUE_DATE", GetType(Date))
                dt1.Columns.Add("OVERRIDE_DUE_DATE", GetType(Date))
                dt1.Columns.Add("LCE_STATUS", GetType(Integer))
                dt1.Columns.Add("STATUS", GetType(String))
                dt1.Columns.Add("OVERRIDE_AMOUNT", GetType(Integer))
                dt1.Columns.Add("POLICY_AMOUNT", GetType(Integer))
                dt1.Columns.Add("SETTLEMENT_AMOUNT", GetType(Integer))
                dt1.Columns.Add("PAID_AMOUNT", GetType(Integer))
                dt1.Columns.Add("DATE_RECEIVED", GetType(Date))
                dt1.Columns.Add("WORKSHOP_DATE", GetType(Date))
                dt1.Columns.Add("WORKSHOP_RESULT", GetType(String))
                dt1.Columns.Add("SHOW_CAUSE_DATE", GetType(Date))
                dt1.Columns.Add("COMMISSION_RESULTS", GetType(String))
                dt1.Columns.Add("COMMISSION_DATE", GetType(Date))
                dt1.Columns.Add("SHOW_CAUSE_RESULTS", GetType(String))
                dt1.Columns.Add("PENDING_LETTER", GetType(String))
                dt1.Columns.Add("LETTER_PRINTED", GetType(Boolean))
                dt1.Columns.Add("LETTER_GENERATED", GetType(Date))
                dt1.Columns.Add("FILLER1", GetType(String))
                dt1.Columns.Add("FILLER2", GetType(String))
                dt1.Columns.Add("FILLER3", GetType(String))
                dt1.Columns.Add("FILLER4", GetType(String))
                dt1.Columns.Add("FILLER5", GetType(String))
                dt1.Columns.Add("Escalation", GetType(String))
                dt1.Columns.Add("PENDING_LETTER_TEMPLATE_NUM", GetType(Integer))

                dt1.Columns("LCE_DATE").ColumnName = "LCE" + vbCrLf + "DATE"
                dt1.Columns("LCE_PROCESS_DATE").ColumnName = "Last" + vbCrLf + "Process Date"
                dt1.Columns("NEXT_DUE_DATE").ColumnName = "Next" + vbCrLf + "Due Date"
                dt1.Columns("OVERRIDE_DUE_DATE").ColumnName = "Override" + vbCrLf + "Due Date"
                dt1.Columns("LCE_STATUS").ColumnName = "Status"
                dt1.Columns("STATUS").ColumnName = "LCE_Status"
                dt1.Columns("OVERRIDE_AMOUNT").ColumnName = "Override" + vbCrLf + "Amount"
                dt1.Columns("POLICY_AMOUNT").ColumnName = "Policy" + vbCrLf + "Amount"
                dt1.Columns("SETTLEMENT_AMOUNT").ColumnName = "Settlement" + vbCrLf + "Amount"
                dt1.Columns("PAID_AMOUNT").ColumnName = "Paid" + vbCrLf + "Amount"
                dt1.Columns("DATE_RECEIVED").ColumnName = "Date" + vbCrLf + "Received"
                dt1.Columns("WORKSHOP_DATE").ColumnName = "WorkShop" + vbCrLf + "Date"
                dt1.Columns("WORKSHOP_RESULT").ColumnName = "WorkShop" + vbCrLf + "Result"
                dt1.Columns("SHOW_CAUSE_DATE").ColumnName = "Show Cause" + vbCrLf + "Hearing Date"
                dt1.Columns("COMMISSION_RESULTS").ColumnName = "Commission" + vbCrLf + "Hearing Results"
                dt1.Columns("COMMISSION_DATE").ColumnName = "Commission" + vbCrLf + "Hearing Date"
                dt1.Columns("SHOW_CAUSE_RESULTS").ColumnName = "Show Cause" + vbCrLf + "Hearing Results"
                dt1.Columns("PENDING_LETTER").ColumnName = "Pending" + vbCrLf + "Letter"
                dt1.Columns("LETTER_PRINTED").ColumnName = "Letter" + vbCrLf + "Printed"
                dt1.Columns("LETTER_GENERATED").ColumnName = "Letter" + vbCrLf + "Generated"


                dt2 = New DataTable
                dt2.Columns.Add("Facility_ID".ToUpper, GetType(Integer))
                dt2.Columns.Add("Facility", GetType(String))
                dt2.Columns.Add("Citation", GetType(Integer))
                dt2.Columns.Add("Due", GetType(Date))
                dt2.Columns.Add("Received", GetType(Date))
                dt2.Columns.Add("Citation Text", GetType(String))

                For Each oLCEInfoLocal In colLCE.Values
                    If (bolLCE = True And bolEnforcementhistory = False) Then
                        If Not (oLCEInfoLocal.PendingLetter = 0 And oLCEInfoLocal.Status = "NFA") Then
                            bolLCEExists = True
                        Else
                            bolLCEExists = False
                        End If
                    ElseIf (bolLCE = False And bolEnforcementhistory = True) Then
                        If (oLCEInfoLocal.PendingLetter = 0 And oLCEInfoLocal.Status = "NFA") Then
                            bolEnfHistoryExists = True
                        Else
                            bolEnfHistoryExists = False
                        End If
                    End If
                    If (bolLCEExists = True) Or (bolEnfHistoryExists = True) Then
                        If oLCEInfoLocal.ID > 0 Then
                            dr = dt.NewRow()
                            dr("LCE_Status") = oLCEInfoLocal.Status
                            dt.Rows.Add(dr)

                            dr1 = dt1.NewRow()
                            dr1("LCE_ID") = oLCEInfoLocal.ID
                            dr1("Licensee") = oLCEInfoLocal.LicenseeName
                            dr1("Licensee_id") = oLCEInfoLocal.LicenseeID
                            dr1("Facility_ID") = oLCEInfoLocal.FacilityID
                            dr1("Rescinded") = oLCEInfoLocal.Rescinded
                            If Date.Compare(oLCEInfoLocal.LCEDate, CDate("01/01/0001")) = 0 Then
                                dr1("LCE" + vbCrLf + "DATE") = DBNull.Value
                            Else
                                dr1("LCE" + vbCrLf + "DATE") = oLCEInfoLocal.LCEDate
                            End If

                            If Date.Compare(oLCEInfoLocal.LCEProcessDate, CDate("01/01/0001")) = 0 Then
                                dr1("Last" + vbCrLf + "Process Date") = DBNull.Value
                            Else
                                dr1("Last" + vbCrLf + "Process Date") = oLCEInfoLocal.LCEProcessDate
                            End If

                            If Date.Compare(oLCEInfoLocal.NextDueDate, CDate("01/01/0001")) = 0 Then
                                dr1("Next" + vbCrLf + "Due Date") = DBNull.Value
                            Else
                                dr1("Next" + vbCrLf + "Due Date") = oLCEInfoLocal.NextDueDate
                            End If

                            If Date.Compare(oLCEInfoLocal.OverrideDueDate, CDate("01/01/0001")) = 0 Then
                                dr1("Override" + vbCrLf + "Due Date") = DBNull.Value
                            Else
                                dr1("Override" + vbCrLf + "Due Date") = oLCEInfoLocal.OverrideDueDate
                            End If

                            dr1("Status") = oLCEInfoLocal.LCEStatus
                            dr1("LCE_Status") = oLCEInfoLocal.Status
                            dr1("Escalation") = GetEscalation(oLCEInfoLocal)
                            'If oLCEInfoLocal.EscalationName Is String.Empty Then
                            '    dr1("Escalation") = oLCEInfoLocal.Status
                            'Else
                            '    dr1("Escalation") = oLCEInfoLocal.EscalationName
                            'End If

                            dr1("Override" + vbCrLf + "Amount") = oLCEInfoLocal.OverrideAmount
                            dr1("Policy" + vbCrLf + "Amount") = oLCEInfoLocal.PolicyAmount
                            dr1("Settlement" + vbCrLf + "Amount") = oLCEInfoLocal.SettlementAmount
                            dr1("Paid" + vbCrLf + "Amount") = oLCEInfoLocal.PaidAmount

                            If Date.Compare(oLCEInfoLocal.DateReceived, CDate("01/01/0001")) = 0 Then
                                dr1("Date" + vbCrLf + "Received") = DBNull.Value
                            Else
                                dr1("Date" + vbCrLf + "Received") = oLCEInfoLocal.DateReceived
                            End If

                            If Date.Compare(oLCEInfoLocal.WorkShopDate, CDate("01/01/0001")) = 0 Then
                                dr1("WorkShop" + vbCrLf + "Date") = DBNull.Value
                            Else
                                dr1("WorkShop" + vbCrLf + "Date") = oLCEInfoLocal.WorkShopDate
                            End If

                            dr1("WorkShop" + vbCrLf + "Result") = oLCEInfoLocal.WorkshopResultName

                            If Date.Compare(oLCEInfoLocal.ShowCauseDate, CDate("01/01/0001")) = 0 Then
                                dr1("Show Cause" + vbCrLf + "Hearing Date") = DBNull.Value
                            Else
                                dr1("Show Cause" + vbCrLf + "Hearing Date") = oLCEInfoLocal.ShowCauseDate
                            End If

                            dr1("Commission" + vbCrLf + "Hearing Results") = oLCEInfoLocal.CommissionResults

                            If Date.Compare(oLCEInfoLocal.CommissionDate, CDate("01/01/0001")) = 0 Then
                                dr1("Commission" + vbCrLf + "Hearing Date") = DBNull.Value
                            Else
                                dr1("Commission" + vbCrLf + "Hearing Date") = oLCEInfoLocal.CommissionDate
                            End If

                            dr1("Show Cause" + vbCrLf + "Hearing Results") = oLCEInfoLocal.ShowCauseResults
                            If oLCEInfoLocal.PendingLetter > 0 Then
                                oProperty.Retrieve(oLCEInfoLocal.PendingLetter)
                                oLCEInfoLocal.PendingLetterName = oProperty.Name
                            Else
                                oLCEInfoLocal.PendingLetterName = String.Empty
                            End If
                            If oLCEInfoLocal.PendingLetter = 1173 Then
                                dr1("Pending" + vbCrLf + "Letter") = "Licensee " + oLCEInfoLocal.PendingLetterName
                            Else
                                dr1("Pending" + vbCrLf + "Letter") = oLCEInfoLocal.PendingLetterName
                            End If
                            dr1("PENDING_LETTER_TEMPLATE_NUM") = oLCEInfoLocal.PendingLetterTemplateNum

                            'dr1("Pending" + vbCrLf + "Letter") = oLCEInfoLocal.PendingLetterName
                            dr1("Letter" + vbCrLf + "Printed") = oLCEInfoLocal.LetterPrinted

                            If Date.Compare(oLCEInfoLocal.LetterGenerated, CDate("01/01/0001")) = 0 Then
                                dr1("Letter" + vbCrLf + "Generated") = DBNull.Value
                            Else
                                dr1("Letter" + vbCrLf + "Generated") = oLCEInfoLocal.LetterGenerated
                            End If

                            dr1("FILLER1") = ""
                            dr1("FILLER2") = ""
                            dr1("FILLER3") = ""
                            dr1("FILLER4") = ""
                            dr1("FILLER5") = ""
                            dt1.Rows.Add(dr1)

                            dr2 = dt2.NewRow
                            dr2("FACILITY_ID") = oLCEInfoLocal.FacilityID
                            dr2("Facility") = oLCEInfoLocal.facilityName
                            dr2("Citation") = oLCEInfoLocal.LicenseeCitationID

                            If Date.Compare(oLCEInfoLocal.CitationDueDate, CDate("01/01/0001")) = 0 Then
                                dr2("Due") = DBNull.Value
                            Else
                                dr2("Due") = oLCEInfoLocal.CitationDueDate
                            End If

                            If Date.Compare(oLCEInfoLocal.CitationReceivedDate, CDate("01/01/0001")) = 0 Then
                                dr2("Received") = DBNull.Value
                            Else
                                dr2("Received") = oLCEInfoLocal.CitationReceivedDate
                            End If
                            dr2("Citation Text") = oLCEInfoLocal.CitationText
                            dt2.Rows.Add(dr2)
                        End If
                    End If
                Next
                Dim dt4 As DataTable
                dt4 = SelectDistinct("dt4", dt, "LCE_Status")

                If Not colLCE.Count = 0 Then
                    ds.Tables.Add(dt4)
                    ds.Tables.Add(dt1)
                    ds.Tables.Add(dt2)

                    dsRel1 = New DataRelation("StatustoLicensee", ds.Tables(0).Columns("LCE_Status"), ds.Tables(1).Columns("LCE_Status"), False)
                    dsrel2 = New DataRelation("LicenseetoFacility", ds.Tables(1).Columns("Facility_ID"), ds.Tables(2).Columns("Facility_ID"), False)

                    ds.Relations.Add(dsRel1)
                    ds.Relations.Add(dsrel2)
                End If
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#Region "Select Distinct Functions"
        Public Function SelectDistinct(ByVal TableName As String, _
                                   ByVal SourceTable As DataTable, _
                                   ByVal FieldName As String) As DataTable
            Dim dt As New DataTable(TableName)
            dt.Columns.Add(FieldName, SourceTable.Columns(FieldName).DataType)
            Dim dr As DataRow, LastValue As Object
            For Each dr In SourceTable.Select("", FieldName)
                If LastValue Is Nothing OrElse Not ColumnEqual(LastValue, dr(FieldName)) Then
                    LastValue = dr(FieldName)
                    dt.Rows.Add(New Object() {LastValue})
                End If
            Next
            Return dt
        End Function
        Private Function ColumnEqual(ByVal A As Object, ByVal B As Object) As Boolean
            '
            ' Compares two values to determine if they are equal. Also compares DBNULL.Value.
            '
            ' NOTE: If your DataTable contains object fields, you must extend this
            ' function to handle them in a meaningful way if you intend to group on them.
            '
            If A Is DBNull.Value And B Is DBNull.Value Then Return True ' Both are DBNull.Value.
            If A Is DBNull.Value Or B Is DBNull.Value Then Return False ' Only one is DbNull.Value.
            Return A = B                                                ' Value type standard comparison
        End Function

#End Region

#End Region
#Region "Lookup Operations"
        Public Function PopulateFacilityName(Optional ByVal OwnerID As Int32 = 0) As DataTable
            Try
                Dim ds As DataSet = oLCEDB.DBGetDS("select * from dbo.vCAE_FacilityName where owner_id = " + OwnerID.ToString)
                Return ds.Tables(0)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateOwnerName() As DataTable
            Try
                Dim ds As DataSet = oLCEDB.DBGetDS("select * from v_OWNER_NAME")
                Return ds.Tables(0)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateCitationList() As DataSet
            Try
                Dim dsReturn As DataSet = oLCEDB.GetCitationList(True)
                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        Public Function getLCELicensees() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VCAE_LCE_Licensees")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function getLicenseeCitation(ByVal LCEID As Integer) As DataTable
            Try
                Return oLCEDB.GetCitation(LCEID).Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function getDropDownValues(ByVal propertyTypeID As Integer) As DataTable
            Try
                Return oLCEDB.GetDropDownValues(propertyTypeID).Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Private Function GetDataTable(ByVal DBViewName As String, Optional ByVal OwnerID As Int32 = 0) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM " & DBViewName
                dsReturn = oLCEDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                Else
                    dtReturn = Nothing
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' This function is used to populate the Inspectors combo box on AssignedInpection UI
        Public Function getInspectors() As DataTable
            Dim dtTable As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCAEInspectors")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function getInspectionTypes() As DataTable
            Dim dtTable As DataTable
            Try
                Dim dtReturn As DataTable = oLCEDB.DBGetDS("select PROPERTY_NAME,PROPERTY_ID from tblsys_property_master where property_type_id = 127").Tables(0)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function getFacilityAssignedInspector(ByVal facilityID As Integer) As Integer
            Try
                Return oLCEDB.getFacilityAssignedInspector(facilityID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub LicenseeComplianceEventInfoChanged(ByVal bolValue As Boolean) Handles oLCEInfo.LCEInfoChanged
            RaiseEvent evtLCEInfoChanged(bolValue)
        End Sub
        Private Sub LicenseeComplianceEventColChanged(ByVal bolValue As Boolean) Handles colLCE.LCEColChanged
            RaiseEvent evtLCECollectionChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
