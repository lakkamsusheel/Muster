'-------------------------------------------------------------------------------
' MUSTER.Info.ZipCodeInfo
'   Provides the container to persist MUSTER Owner state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN       12/23/04    Original class definition.
'  1.1        AN       12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        AB       02/22/05    Added AgeThreshold and IsAgedData Attributes
'
'
' Function          Description
' New()             Instantiates an empty Zip object
' New(Zip,state,City,County,Fips,CREATED_BY,DATE_CREATED,LAST_EDITED_BY,DATE_LAST_EDITED)
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
' Archive            Sets the object state to the old state when loaded from or
'                   last saved to the repository
' CheckDirty         'Check for dirty....
' Init               Intialise the object attributes...

' Attribute          Description
'ID  Unique identifier(ID:=(Me.strZIP & "|" & Me.strState & "|" & Me.strCity & "|" & Me.strCounty))for the Collection object
'Zip ZipCode..
'City City.
'State State
'County County
'FIPS Fips
'IsDirty - To check the object is Dirty or not..

'Read Only Attribute...

'CreatedBy - User name creating the record
'CreatedOn  - Created date
'ModifiedBy - User name modifying the record
'ModifiedOn - Modified Date..
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class ZipCodeInfo
        Implements iAccessors
#Region "Private member variables"

        'Private nZIPID As Integer
        Private strZIP As String
        Private strState As String
        Private strCity As String
        Private strCounty As String
        Private strFips As String
        Private strCreatedBy As String
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private bolIsDirty As Boolean

        'Current values
        ' Private onZIPID As Integer
        Private ostrZIP As String
        Private ostrState As String
        Private ostrCity As String
        Private ostrCounty As String
        Private ostrFips As String
        Private ostrCreatedBy As String
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private obolIsDirty As Boolean
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
            Me.Init()
        End Sub
        ' The prototype New method
        Public Sub New(ByVal Zip As String, ByVal state As String, ByVal City As String, ByVal County As String, ByVal Fips As String, ByVal CREATED_BY As String, ByVal DATE_CREATED As Date, ByVal LAST_EDITED_BY As String, ByVal DATE_LAST_EDITED As Date)
            ostrZIP = Zip
            ostrState = state
            ostrCity = City
            ostrCounty = County
            ostrFips = Fips
            ostrCreatedBy = CREATED_BY
            odtCreatedOn = DATE_CREATED
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()

            strZIP = ostrZIP
            strState = ostrState
            strCity = ostrCity
            strCounty = ostrCounty
            strFips = ostrFips
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

        End Sub
        Public Sub Archive()

            ostrZIP = strZIP
            ostrState = strState
            ostrCity = strCity
            ostrCounty = strCounty
            ostrFips = strFips
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            bolIsDirty = (ostrZIP <> strZIP) Or _
                 (ostrState <> strState) Or _
                 (ostrCity <> strCity) Or _
                 (ostrCounty <> strCounty) Or _
                 (ostrFips <> strFips) Or _
                 (ostrCreatedBy <> strCreatedBy) Or _
                 (odtCreatedOn <> dtCreatedOn) Or _
                (ostrModifiedBy <> strModifiedBy) Or _
                (odtModifiedOn <> dtModifiedOn)
        End Sub
        Private Sub Init()
            ostrZIP = String.Empty
            ostrState = String.Empty
            ostrCity = String.Empty
            ostrCounty = String.Empty
            ostrFips = String.Empty
            ostrCreatedBy = String.Empty
            odtCreatedOn = System.DateTime.Now
            ostrModifiedBy = String.Empty
            odtModifiedOn = System.DateTime.Now
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return Me.strZIP & "|" & Me.strState & "|" & Me.strCity & "|" & Me.strCounty
            End Get
            Set(ByVal value As String)
                Try
                    Dim arrVals() As String
                    arrVals = value.Split("|")
                    strZIP = arrVals(0)
                    strCity = arrVals(1)
                    strState = arrVals(2)
                    strCounty = arrVals(3)
                Catch Ex As Exception
                    MusterException.Publish(Ex, Nothing, Nothing)
                    Throw Ex
                End Try
                Me.CheckDirty()
            End Set
        End Property
        Public Property Zip() As String
            Get
                Return Me.strZIP
            End Get
            Set(ByVal value As String)
                Me.strZIP = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property City() As String
            Get
                Return Me.strCity
            End Get
            Set(ByVal value As String)
                Me.strCity = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property County() As String
            Get
                Return Me.strCounty
            End Get
            Set(ByVal value As String)
                Me.strCounty = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property state() As String
            Get
                Return Me.strState
            End Get
            Set(ByVal value As String)
                Me.strState = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Fips() As String
            Get
                Return Me.strFips
            End Get
            Set(ByVal value As String)
                Me.strFips = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property AgeThreshold() As Int16
            Get
                Return nAgeThreshold
            End Get

            Set(ByVal value As Int16)
                nAgeThreshold = Int16.Parse(value)
            End Set
        End Property

        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property

#End Region
#Region "iAccessors"
        Public ReadOnly Property CreatedBy() As String Implements iAccessors.CreatedBy
            Get
                Return strCreatedBy
            End Get
        End Property
        Public ReadOnly Property CreatedOn() As Date Implements iAccessors.CreatedOn
            Get
                Return dtCreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedBy() As String Implements iAccessors.ModifiedBy
            Get
                Return strModifiedBy
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date Implements iAccessors.ModifiedOn
            Get
                Return dtModifiedOn
            End Get
        End Property
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
