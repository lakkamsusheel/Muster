'-------------------------------------------------------------------------------
' MUSTER.Info.PersonaInfo.vb
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN     12/13/04    Original class definition.
'  1.1        AN     12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        EN     01/10/05    Added Archive method. 
'  1.3        EN     01/14/05     modified the variable strorganization_Entity_code  to norganization_Entity_code
'  1.3        MNR    01/13/05    Added Events
'  1.4        JVCII  02/11/05    Modified Reset to check for presence of pipe
'                                   delimiter when reseting strID.  Still not certain
'                                   what strID is actually used for...
'                                Added new property Ambi_ID which returns either
'                                   nPersonID or nOrgID or 0.
'                                Modified PersonID and OrgID to set the "other"
'                                   ID value to 0 when either is set to a non-zero
'                                   value.
'  1.5        AB     02/18/05    Added AgeThreshold and IsAgedData Attributes
'  1.6        MNR    03/16/05    Removed strSrc from events
'  1.7        MNR    03/22/05    Added Constructor New(ByVal drPersona As DataRow, ByVal personaType As String)
'
' Function          Description
'  New()             Instantiates an empty PersonaInfo object.
'  New(ID, PersonID, OrgID, organization_Entity_code, Company, Title,
'       Prefix, Firstname, Middlename, Lastname, Suffix,
'       Deleted,CreatedBy, CreatedOn, ModifiedBy, LastEdited)
'                   Instantiates a populated OwnerInfo object
'' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'  
'Attribute          Description
'  ID         The unique identifier associated with the Person in the repository.
'  PersonID   The unique identifier for person records.
'  OrgID      The unique identifier for Org records.
'  organization_Entity_code Entity code for Org records
'  Company   Company name
'  Title      Title For person record.
'  Prefix     Prefix for Person record.
'  Firstname  First name for person record.
'  Middlename Middle name for person record.
'  Lastname  last Name for person record.
'  Suffix    suffix for person record.
'  Deleted   deleted flag indicating record is deleted from the table...
'  IsDirty   Indicates if the Address state has been altered since it was
'                       last loaded from or saved to the repository.
Namespace MUSTER.Info
    <Serializable()> _
Public Class PersonaInfo

#Region "Public Events"
        Public Event PersonaInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        'Original Values

        Private nPersonId As Integer
        '        Private bolShowDeleted As Boolean
        Private strID As String
        Private nOrgId As Integer
        Private strCompanyname As String
        Private norganization_Entity_code As Integer
        Private bolDELETED As Boolean
        Private strTitle As String
        Private strPrefix As String
        Private strFirstname As String
        Private strMiddlename As String = Nothing
        Private strLastname As String
        Private strSuffix As String
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString

        'Current Values
        Private ostrID As String
        Private ostrTitle As String
        Private ostrPrefix As String
        Private ostrFirstname As String
        Private ostrMiddlename As String = Nothing
        Private ostrLastname As String
        Private ostrSuffix As String
        Private ostrCompanyname As String
        Private onOrgId As Integer
        Private onPersonId As Integer
        Private obolDELETED As Boolean
        Private onorganization_Entity_code As Integer

        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5

#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
            Me.Init()
        End Sub
        Sub New(ByVal ID As String, ByVal PersonID As Integer, _
            ByVal OrgID As Integer, _
            ByVal organization_Entity_code As Integer, _
            ByVal Company As String, _
            ByVal Title As String, _
            ByVal Prefix As String, _
            ByVal Firstname As String, _
            ByVal Middlename As String, _
            ByVal Lastname As String, _
            ByVal Suffix As String, _
            ByVal Deleted As Boolean, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date)
            ostrID = ID
            onPersonId = PersonID
            onOrgId = OrgID
            onorganization_Entity_code = organization_Entity_code
            ostrCompanyname = Company
            ostrTitle = Title
            ostrPrefix = Prefix
            ostrFirstname = Firstname
            ostrMiddlename = Middlename
            ostrLastname = Lastname
            ostrSuffix = Suffix
            obolDELETED = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            dtDataAge = Now()
            Me.Reset()
        End Sub
        Sub New(ByVal drPersona As DataRow, ByVal personaType As String)
            Select Case UCase(personaType).Trim
                Case "O"
                    onOrgId = drPersona.Item("ORGANIZATION_ID")
                    ostrID = "O|" + onOrgId.ToString()
                    onorganization_Entity_code = IIf(drPersona.Item("ORGANIZATION_ENTITY_CODE") Is System.DBNull.Value, 0, drPersona.Item("ORGANIZATION_ENTITY_CODE"))
                    ostrCompanyname = drPersona.Item("NAME")
                    obolDELETED = drPersona.Item("DELETED")
                    ostrCreatedBy = drPersona.Item("CREATED_BY")
                    odtCreatedOn = IIf(drPersona.Item("DATE_CREATED") Is System.DBNull.Value, CDate("01/01/0001"), drPersona.Item("DATE_CREATED"))
                    ostrModifiedBy = IIf(drPersona.Item("LAST_EDITED_BY") Is System.DBNull.Value, String.Empty, drPersona.Item("LAST_EDITED_BY"))
                    odtModifiedOn = IIf(drPersona.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), drPersona.Item("DATE_LAST_EDITED"))
                Case "P"
                    onPersonId = drPersona.Item("PERSON_ID")
                    ostrID = "P|" + onPersonId.ToString()
                    ostrTitle = IIf(drPersona.Item("TITLE") Is System.DBNull.Value, String.Empty, drPersona.Item("TITLE"))
                    ostrPrefix = IIf(drPersona.Item("PREFIX") Is System.DBNull.Value, String.Empty, drPersona.Item("PREFIX"))
                    ostrFirstname = drPersona.Item("FIRST_NAME")
                    ostrMiddlename = IIf(drPersona.Item("MIDDLE_NAME") Is System.DBNull.Value, String.Empty, drPersona.Item("MIDDLE_NAME"))
                    ostrLastname = drPersona.Item("LAST_NAME")
                    ostrSuffix = IIf(drPersona.Item("SUFFIX") Is System.DBNull.Value, String.Empty, drPersona.Item("SUFFIX"))
                    obolDELETED = drPersona.Item("DELETED")
                    ostrCreatedBy = drPersona.Item("CREATED_BY")
                    odtCreatedOn = IIf(drPersona.Item("DATE_CREATED") Is System.DBNull.Value, CDate("01/01/0001"), drPersona.Item("DATE_CREATED"))
                    ostrModifiedBy = IIf(drPersona.Item("LAST_EDITED_BY") Is System.DBNull.Value, String.Empty, drPersona.Item("LAST_EDITED_BY"))
                    odtModifiedOn = IIf(drPersona.Item("DATE_LAST_EDITED") Is System.DBNull.Value, CDate("01/01/0001"), drPersona.Item("DATE_LAST_EDITED"))
            End Select
            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If strID Is Nothing Then
                strID = ostrID
            Else
                If strID.IndexOf("|") < 0 Then
                    If IsNumeric(strID) Then
                        If Integer.Parse(strID) > 0 Then
                            strID = ostrID
                        End If
                    End If
                Else
                    Dim keyID As Integer = CType(strID.Split("|")(1), Integer)
                    If keyID >= 0 Then
                        strID = ostrID
                    End If
                End If
            End If
            nPersonId = onPersonId
            nOrgId = onOrgId
            norganization_Entity_code = onorganization_Entity_code
            strCompanyname = ostrCompanyname
            strTitle = ostrTitle
            strPrefix = ostrPrefix
            strFirstname = ostrFirstname
            strMiddlename = ostrMiddlename
            strLastname = ostrLastname
            strSuffix = ostrSuffix
            bolDELETED = obolDELETED
            bolIsDirty = False
            RaiseEvent PersonaInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            ostrID = strID
            onPersonId = nPersonId
            onOrgId = nOrgId
            onorganization_Entity_code = norganization_Entity_code
            ostrCompanyname = strCompanyname
            ostrTitle = strTitle
            ostrPrefix = strPrefix
            ostrFirstname = strFirstname
            ostrMiddlename = strMiddlename
            ostrLastname = strLastname
            ostrSuffix = strSuffix
            obolDELETED = bolDELETED
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = False
            obolIsDirty = (nOrgId <> onOrgId Or _
                       norganization_Entity_code <> onorganization_Entity_code Or _
                        strCompanyname <> ostrCompanyname Or _
                        strTitle <> ostrTitle Or _
                        strPrefix <> ostrPrefix Or _
                        strFirstname <> ostrFirstname Or _
                        strMiddlename <> ostrMiddlename Or _
                        strLastname <> ostrLastname Or _
                        strSuffix <> ostrSuffix Or _
                        bolDELETED <> obolDELETED)
            If obolIsDirty Then
                RaiseEvent PersonaInfoChanged(obolIsDirty)
            End If
        End Sub
        Private Sub Init()
            ostrTitle = String.Empty
            ostrPrefix = String.Empty
            ostrFirstname = String.Empty
            ostrMiddlename = String.Empty
            ostrLastname = String.Empty
            ostrSuffix = String.Empty
            ostrCompanyname = String.Empty
            onorganization_Entity_code = 0
            onOrgId = 0
            onPersonId = 0
            ostrID = String.Empty
            obolDELETED = False
            dtCreatedOn = System.DateTime.Now
            dtModifiedOn = System.DateTime.Now
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"

        Public ReadOnly Property Ambi_ID() As Integer
            Get
                '
                ' Returns either the non-zero Org or Person ID
                '   or zero if both are zero.
                If nPersonId <> 0 Then
                    Return nPersonId
                ElseIf nOrgId <> 0 Then
                    Return nOrgId
                Else
                    Return 0
                End If
            End Get
        End Property
        Public Property ID() As String
            Get
                Return strID
            End Get
            Set(ByVal Value As String)
                strID = Value
                Me.CheckDirty()
            End Set
        End Property


        Public Property PersonId() As Integer
            Get
                Return nPersonId
            End Get

            Set(ByVal value As Integer)
                '
                ' WTF??? - Why cast and integer to an integer?
                '                nPersonId = Integer.Parse(value)
                nPersonId = value
                '
                ' Also need to reset nOrgID
                '
                If value <> 0 Then nOrgId = 0
                Me.CheckDirty()
            End Set
        End Property
        Public Property OrgID() As Integer
            Get
                Return nOrgId
            End Get

            Set(ByVal value As Integer)
                '
                ' WTF??? - Why cast and integer to an integer?
                '                nOrgId = Integer.Parse(value)
                nOrgId = value
                '
                ' Also need to reset nPersonID
                '
                If value <> 0 Then nPersonId = 0
                Me.CheckDirty()
            End Set
        End Property
        Public Property Org_Entity_Code() As Integer
            Get
                Return norganization_Entity_code
            End Get

            Set(ByVal value As Integer)
                norganization_Entity_code = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Company() As String
            Get
                Return strCompanyname
            End Get

            Set(ByVal value As String)
                strCompanyname = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Title() As String
            Get
                Return strTitle
            End Get

            Set(ByVal value As String)
                strTitle = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Prefix() As String
            Get
                Return strPrefix
            End Get

            Set(ByVal value As String)
                strPrefix = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FirstName() As String
            Get
                Return strFirstname
            End Get

            Set(ByVal value As String)
                strFirstname = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property MiddleName() As String
            Get
                Return strMiddlename
            End Get

            Set(ByVal value As String)
                strMiddlename = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LastName() As String
            Get
                Return strLastname
            End Get

            Set(ByVal value As String)
                strLastname = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Suffix() As String
            Get
                Return strSuffix
            End Get

            Set(ByVal value As String)
                strSuffix = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDELETED
            End Get

            Set(ByVal value As Boolean)
                bolDELETED = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
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


        'Public Property ShowDeleted() As Boolean
        '    Get
        '        Return bolShowDeleted
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        bolShowDeleted = Value
        '    End Set
        'End Property

#Region "iAccessors"
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property
#End Region
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class

End Namespace


