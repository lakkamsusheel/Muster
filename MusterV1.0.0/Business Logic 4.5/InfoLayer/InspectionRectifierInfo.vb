'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionRectifierInfo
'   Provides the container to persist MUSTER InspectionRectifier state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionRectifierInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionRectifierInfo object
' New(dr)           Instantiates a populated InspectionRectifierInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionRectifier to build other objects.
'       Replace keyword "InspectionRectifier" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionRectifierInfo
#Region "Public Events"
        Public Event evtInspectionRectifierInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsRectID As Int64
        Private nInspectionID As Int64
        Private nQuestionID As Int64
        Private bolRectifierOn As Boolean
        Private strInopHowLong As String
        Private dVolts As Double
        Private dAmps As Double
        Private dHours As Double
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime

        Private onInsRectID As Int64
        Private onInspectionID As Int64
        Private onQuestionID As Int64
        Private obolRectifierOn As Boolean
        Private ostrInopHowLong As String
        Private odVolts As Double
        Private odAmps As Double
        Private odHours As Double
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal id As Int64, _
        ByVal inspectionID As Int64, _
        ByVal questionID As Int64, _
        ByVal rectifierOn As Boolean, _
        ByVal inopHowLong As String, _
        ByVal volts As Double, _
        ByVal amps As Double, _
        ByVal hours As Double, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As Date, _
        ByVal modifiedBy As String, _
        ByVal modifiedOn As Date)
            onInsRectID = id
            onInspectionID = inspectionID
            onQuestionID = questionID
            obolRectifierOn = rectifierOn
            ostrInopHowLong = inopHowLong
            odVolts = volts
            odAmps = amps
            odHours = hours
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedOn
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionRectifier As DataRow)
            Try
                onInsRectID = drInspectionRectifier.Item("INS_RECT_ID")
                onInspectionID = IIf(drInspectionRectifier.Item("INSPECTION_ID") Is DBNull.Value, 0, drInspectionRectifier.Item("INSPECTION_ID"))
                onQuestionID = IIf(drInspectionRectifier.Item("QUESTION_ID") Is DBNull.Value, 0, drInspectionRectifier.Item("QUESTION_ID"))
                obolRectifierOn = IIf(drInspectionRectifier.Item("RECITIFIER_ON") Is DBNull.Value, False, drInspectionRectifier.Item("RECITIFIER_ON"))
                ostrInopHowLong = IIf(drInspectionRectifier.Item("INOP_HOW_LONG") Is DBNull.Value, String.Empty, drInspectionRectifier.Item("INOP_HOW_LONG"))
                odVolts = IIf(drInspectionRectifier.Item("VOLTS") Is DBNull.Value, 0.0, drInspectionRectifier.Item("VOLTS"))
                odAmps = IIf(drInspectionRectifier.Item("AMPS") Is DBNull.Value, 0.0, drInspectionRectifier.Item("AMPS"))
                odHours = IIf(drInspectionRectifier.Item("HOURS") Is DBNull.Value, 0.0, drInspectionRectifier.Item("HOURS"))
                obolDeleted = IIf(drInspectionRectifier.Item("DELETED") Is DBNull.Value, False, drInspectionRectifier.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionRectifier.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionRectifier.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionRectifier.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionRectifier.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionRectifier.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionRectifier.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionRectifier.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionRectifier.Item("DATE_LAST_EDITED"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nInsRectID >= 0 Then
                nInsRectID = onInsRectID
            End If
            nInspectionID = onInspectionID
            nQuestionID = onQuestionID
            bolRectifierOn = obolRectifierOn
            strInopHowLong = ostrInopHowLong
            dVolts = odVolts
            dAmps = odAmps
            dHours = odHours
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent evtInspectionRectifierInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInsRectID = nInsRectID
            onInspectionID = nInspectionID
            onQuestionID = nQuestionID
            obolRectifierOn = bolRectifierOn
            ostrInopHowLong = strInopHowLong
            odVolts = dVolts
            odAmps = dAmps
            odHours = dHours
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nInspectionID <> onInspectionID) Or _
            (nQuestionID <> onQuestionID) Or _
            (bolRectifierOn <> obolRectifierOn) Or _
            (strInopHowLong <> ostrInopHowLong) Or _
            (dVolts <> odVolts) Or _
            (dAmps <> odAmps) Or _
            (dHours <> odHours) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionRectifierInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInsRectID = 0
            onInspectionID = 0
            onQuestionID = 0
            obolRectifierOn = False
            ostrInopHowLong = String.Empty
            odVolts = 0.0
            odAmps = 0.0
            odHours = 0.0
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nInsRectID
            End Get
            Set(ByVal Value As Int64)
                nInsRectID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InspectionID() As Int64
            Get
                Return nInspectionID
            End Get
            Set(ByVal Value As Int64)
                nInspectionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property QuestionID() As Int64
            Get
                Return nQuestionID
            End Get
            Set(ByVal Value As Int64)
                nQuestionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RectifierOn() As Boolean
            Get
                Return bolRectifierOn
            End Get
            Set(ByVal Value As Boolean)
                bolRectifierOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InopHowLong() As String
            Get
                Return strInopHowLong
            End Get
            Set(ByVal Value As String)
                strInopHowLong = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Volts() As Double
            Get
                Return dVolts
            End Get
            Set(ByVal Value As Double)
                dVolts = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Amps() As Double
            Get
                Return dAmps
            End Get
            Set(ByVal Value As Double)
                dAmps = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Hours() As Double
            Get
                Return dHours
            End Get
            Set(ByVal Value As Double)
                dHours = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal value As Boolean)
                bolDeleted = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
                RaiseEvent evtInspectionRectifierInfoChanged(bolIsDirty)
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        Public Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As Date)
                dtCreatedOn = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
            End Set
        End Property
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
