'-------------------------------------------------------------------------------
' MUSTER.Info.TecDocInfo
'   Provides the container to persist MUSTER Technical Document state
'
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC       05/24/05    Original class definition.
'
' Function          Description
'-------------------------------------------------------------------------------
'

Namespace MUSTER.Info
    Public Class TecDocInfo
#Region "Private Member Variables"
        Private bolActive As Boolean
        Private bolDeleted As Boolean
        Private bolNTFE_Flag As Boolean
        Private bolSTFS_Flag As Boolean
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private nAuto_Doc_1 As Long
        Private nAuto_Doc_2 As Long
        Private nAuto_Doc_3 As Long
        Private nAuto_Doc_4 As Long
        Private nAuto_Doc_5 As Long
        Private nAuto_Doc_6 As Long
        Private nAuto_Doc_7 As Long
        Private nAuto_Doc_8 As Long
        Private nAuto_Doc_9 As Long
        Private nAuto_Doc_10 As Long
        Private nDocType As Long
        Private nEntityID As Integer
        Private obolActive As Boolean
        Private obolNTFE_Flag As Boolean
        Private obolSTFS_Flag As Boolean
        Private obolDeleted As Boolean
        Private onAuto_Doc_1 As Long
        Private onAuto_Doc_2 As Long
        Private onAuto_Doc_3 As Long
        Private onAuto_Doc_4 As Long
        Private onAuto_Doc_5 As Long
        Private onAuto_Doc_6 As Long
        Private onAuto_Doc_7 As Long
        Private onAuto_Doc_8 As Long
        Private onAuto_Doc_9 As Long
        Private onAuto_Doc_10 As Long
        Private onDocType As Long
        Private ostrName As String
        Private ostrPhysical_File_Name As String
        Private onTriggerField As Int64
        Private strName As String
        Private strPhysical_File_Name As String
        Private nTriggerField As Int64
        Private nFinActivityType As Integer
        Private onFinActivityType As Integer


        Private bolIsDirty As Boolean
        Private nDocID As Long
        Private strCreatedBy As String
        Private strModifiedBy As String
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
#End Region
#Region "Public Events"
        Public Delegate Sub InfoChangedEventHandler()
        Public Event InfoChanged As InfoChangedEventHandler
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
        End Sub
        Public Sub New(ByVal Id As Long, _
            ByVal DocType As Long, _
            ByVal sName As String, _
            ByVal sFileName As String, _
            ByVal nTrigger As Int64, _
            ByVal bNTFE As Boolean, _
            ByVal bSTFS As Boolean, _
            ByVal AutoDoc1 As Long, _
            ByVal AutoDoc2 As Long, _
            ByVal AutoDoc3 As Long, _
            ByVal AutoDoc4 As Long, _
            ByVal AutoDoc5 As Long, _
            ByVal AutoDoc6 As Long, _
            ByVal AutoDoc7 As Long, _
            ByVal AutoDoc8 As Long, _
            ByVal AutoDoc9 As Long, _
            ByVal AutoDoc10 As Long, _
            ByVal CreatedBy As String, _
            ByVal CreateDate As Date, _
            ByVal LastEditedBy As String, _
            ByVal LastEditDate As Date, _
            ByVal bActive As Boolean, _
            ByVal bDeleted As Boolean, Optional ByVal finActivityTypeID As Integer = 0)

            nDocID = Id
            onDocType = DocType
            ostrName = sName
            ostrPhysical_File_Name = sFileName
            onTriggerField = nTrigger
            onFinActivityType = finActivityTypeID

            obolNTFE_Flag = bNTFE
            obolSTFS_Flag = bSTFS

            onAuto_Doc_1 = AutoDoc1
            onAuto_Doc_2 = AutoDoc2
            onAuto_Doc_3 = AutoDoc3
            onAuto_Doc_4 = AutoDoc4
            onAuto_Doc_5 = AutoDoc5
            onAuto_Doc_6 = AutoDoc6
            onAuto_Doc_7 = AutoDoc7
            onAuto_Doc_8 = AutoDoc8
            onAuto_Doc_9 = AutoDoc9
            onAuto_Doc_10 = AutoDoc10

            strCreatedBy = CreatedBy
            strModifiedBy = LastEditedBy
            dtModifiedOn = LastEditDate
            dtCreatedOn = CreateDate
            obolActive = bActive
            obolDeleted = bDeleted

            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"

        'Link Tec Doc to Spectific Financial Activity Type
        Public Property FinActivityType() As Integer
            Get
                Return nFinActivityType
            End Get
            Set(ByVal Value As Integer)
                nFinActivityType = Value
            End Set
        End Property

        ' the Active/Inactive flag for the TEC_DOC
        Public Property Active() As Boolean
            Get
                Return bolActive
            End Get
            Set(ByVal Value As Boolean)
                bolActive = Value
                CheckDirty()
            End Set
        End Property
        ' The maximum age the info object can attain before requiring a refresh
        Public Property AgeThreshold() As Date
            Get
                Return dtDataAge
            End Get
            Set(ByVal Value As Date)
                dtDataAge = Value
            End Set
        End Property
        ' The 1st automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_1() As Long
            Get
                Return nAuto_Doc_1
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_1 = Value
                CheckDirty()
            End Set
        End Property
        ' The 2nd automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_2() As Long
            Get
                Return nAuto_Doc_2
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_2 = Value
                CheckDirty()
            End Set
        End Property
        ' The 3rd automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_3() As Long
            Get
                Return nAuto_Doc_3
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_3 = Value
                CheckDirty()
            End Set
        End Property
        ' The 4th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_4() As Long
            Get
                Return nAuto_Doc_4
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_4 = Value
                CheckDirty()
            End Set
        End Property
        ' The 5th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_5() As Long
            Get
                Return nAuto_Doc_5
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_5 = Value
                CheckDirty()
            End Set
        End Property
        ' The 6th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_6() As Long
            Get
                Return nAuto_Doc_6
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_6 = Value
                CheckDirty()
            End Set
        End Property
        ' The 7th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_7() As Long
            Get
                Return nAuto_Doc_7
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_7 = Value
                CheckDirty()
            End Set
        End Property
        ' The 8th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_8() As Long
            Get
                Return nAuto_Doc_8
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_8 = Value
                CheckDirty()
            End Set
        End Property
        ' The 9th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_9() As Long
            Get
                Return nAuto_Doc_9
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_9 = Value
                CheckDirty()
            End Set
        End Property
        ' The 10th automatically generated document associated with the TEC_DOC
        Public Property Auto_Doc_10() As Long
            Get
                Return nAuto_Doc_10
            End Get
            Set(ByVal Value As Long)
                nAuto_Doc_10 = Value
                CheckDirty()
            End Set
        End Property
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
        End Property
        ' The deleted flag for the TEC_DOC
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                CheckDirty()
            End Set
        End Property
        ' The type of document (from tblSYS_PROPERTY_MASTER with property type "TECH DOC")
        Public Property DocType() As Long
            Get
                Return nDocType
            End Get
            Set(ByVal Value As Long)
                nDocType = Value
                CheckDirty()
            End Set
        End Property
        ' The entity ID associated with a technical document.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return nEntityID
            End Get
        End Property
        ' The system ID for this Technical Document
        Public Property ID() As Long
            Get
                Return nDocID
            End Get
            Set(ByVal Value As Long)
                nDocID = Value
            End Set
        End Property
        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
        ' Raised when any of the TechnicalEventInfo attributes are modified
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = False
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
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property
        ' The name of the technical document to be displayed throughout the system
        Public Property Name() As String
            Get
                Return strName
            End Get
            Set(ByVal Value As String)
                strName = Value
                CheckDirty()
            End Set
        End Property
        ' The NTFE-EUD flag for the TEC_DOC
        Public Property NTFE_Flag() As Boolean
            Get
                Return bolNTFE_Flag
            End Get
            Set(ByVal Value As Boolean)
                bolNTFE_Flag = Value
                CheckDirty()
            End Set
        End Property
        ' The physical file name for the Word document associated to the TEC_DOC
        Public Property Physical_File_Name() As String
            Get
                Return strPhysical_File_Name
            End Get
            Set(ByVal Value As String)
                strPhysical_File_Name = Value
                CheckDirty()
            End Set
        End Property
        ' The STFS/STFS-DIRECT/Federal flag for the TEC_DOC
        Public Property STFS_Flag() As Boolean
            Get
                Return bolSTFS_Flag
            End Get
            Set(ByVal Value As Boolean)
                bolSTFS_Flag = Value
                CheckDirty()
            End Set
        End Property
        ' The field in the TechnicalDocument that triggers the automatically generated document
        Public Property Trigger_Field() As Int64
            Get
                Return nTriggerField
            End Get
            Set(ByVal Value As Int64)
                nTriggerField = Value
                CheckDirty()
            End Set
        End Property
#End Region
#Region "Exposed Methods"
        Public Sub Archive()
            obolActive = bolActive
            obolDeleted = bolDeleted
            obolNTFE_Flag = bolNTFE_Flag
            obolSTFS_Flag = bolSTFS_Flag

            onAuto_Doc_1 = nAuto_Doc_1
            onAuto_Doc_2 = nAuto_Doc_2
            onAuto_Doc_3 = nAuto_Doc_3
            onAuto_Doc_4 = nAuto_Doc_4
            onAuto_Doc_5 = nAuto_Doc_5
            onAuto_Doc_6 = nAuto_Doc_6
            onAuto_Doc_7 = nAuto_Doc_7
            onAuto_Doc_8 = nAuto_Doc_8
            onAuto_Doc_9 = nAuto_Doc_9
            onAuto_Doc_10 = nAuto_Doc_10
            onDocType = nDocType
            onFinActivityType = nFinActivityType

            ostrName = strName
            ostrPhysical_File_Name = strPhysical_File_Name
            onTriggerField = nTriggerField
            bolIsDirty = False

        End Sub
        Public Sub Reset()
            bolActive = obolActive
            bolDeleted = obolDeleted
            nFinActivityType = onFinActivityType
            bolNTFE_Flag = obolNTFE_Flag
            bolSTFS_Flag = obolSTFS_Flag

            nAuto_Doc_1 = onAuto_Doc_1
            nAuto_Doc_2 = onAuto_Doc_2
            nAuto_Doc_3 = onAuto_Doc_3
            nAuto_Doc_4 = onAuto_Doc_4
            nAuto_Doc_5 = onAuto_Doc_5
            nAuto_Doc_6 = onAuto_Doc_6
            nAuto_Doc_7 = onAuto_Doc_7
            nAuto_Doc_8 = onAuto_Doc_8
            nAuto_Doc_9 = onAuto_Doc_9
            nAuto_Doc_10 = onAuto_Doc_10
            nDocType = onDocType

            strName = ostrName
            strPhysical_File_Name = ostrPhysical_File_Name
            nTriggerField = onTriggerField

            bolIsDirty = False
        End Sub
#End Region
#Region "Private Methods"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (obolActive <> bolActive) Or _
                        (obolDeleted <> bolDeleted) Or _
                        (obolNTFE_Flag <> bolNTFE_Flag) Or _
                        (obolSTFS_Flag <> bolSTFS_Flag) Or _
                        (onAuto_Doc_1 <> nAuto_Doc_1) Or _
                        (onAuto_Doc_2 <> nAuto_Doc_2) Or _
                        (onFinActivityType <> nFinActivityType) Or _
                        (onAuto_Doc_3 <> nAuto_Doc_3) Or _
                        (onAuto_Doc_4 <> nAuto_Doc_4) Or _
                        (onAuto_Doc_5 <> nAuto_Doc_5) Or _
                        (onAuto_Doc_6 <> nAuto_Doc_6) Or _
                        (onAuto_Doc_7 <> nAuto_Doc_7) Or _
                        (onAuto_Doc_8 <> nAuto_Doc_8) Or _
                        (onAuto_Doc_9 <> nAuto_Doc_9) Or _
                        (onAuto_Doc_10 <> nAuto_Doc_10) Or _
                        (onDocType <> nDocType) Or _
                        (ostrName <> strName) Or _
                        (ostrPhysical_File_Name <> strPhysical_File_Name) Or _
                        (onTriggerField <> nTriggerField)

        End Sub
        Sub Init()

            bolActive = True
            bolDeleted = False
            bolNTFE_Flag = False
            bolSTFS_Flag = False
            dtDataAge = System.DateTime.Now
            nAgeThreshold = 5
            nAuto_Doc_1 = 0
            nAuto_Doc_2 = 0
            nAuto_Doc_3 = 0
            nAuto_Doc_4 = 0
            nAuto_Doc_5 = 0
            nAuto_Doc_6 = 0
            nAuto_Doc_7 = 0
            nAuto_Doc_8 = 0
            nAuto_Doc_9 = 0
            nAuto_Doc_10 = 0
            nFinActivityType = 0

            nDocType = 0
            nEntityID = 0
            obolActive = True
            obolNTFE_Flag = False
            obolSTFS_Flag = False
            obolDeleted = False
            onAuto_Doc_1 = 0
            onAuto_Doc_2 = 0
            onAuto_Doc_3 = 0
            onAuto_Doc_4 = 0
            onDocType = 0
            ostrName = String.Empty
            ostrPhysical_File_Name = String.Empty
            onTriggerField = 0
            strName = String.Empty
            strPhysical_File_Name = String.Empty
            nTriggerField = 0

            bolIsDirty = False
            nDocID = 0
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            dtModifiedOn = DateTime.Now.ToShortDateString
            dtCreatedOn = DateTime.Now.ToShortDateString

        End Sub
#End Region
#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region
    End Class
End Namespace
