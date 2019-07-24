
Namespace MUSTER.Info
    Public Class LustRemediationInfo
#Region "Public Events"
        Public Delegate Sub InfoChangedEventHandler()
        ' Raised when any of the LustRemediationInfo attributes are modified
        Public Event LustRemediationInfoChanged As InfoChangedEventHandler
#End Region
#Region "Private Member Variables"
        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean
        Private nOwned As Integer
        Private nMountType As Integer
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime
        Private dtDatePlacedInUse As Date
        Private dtPurchaseDate As Date
        Private dtRefurbDate As Date
        Private nAgeThreshold As Int16 = 5
        Private nEntityID As Integer
        Private nIDofParent As Long
        Private nOption1 As Long
        Private nOption2 As Long
        Private nOption3 As Long
        Private nRemSysType As Long
        Private obolDeleted As Boolean
        Private obolIsDirty As Boolean
        Private onOwned As Integer
        Private onMountType As Integer
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private odtDatePlacedInUse As Date
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private odtPurchaseDate As Date
        Private odtRefurbDate As Date
        Private onIDofParent As Long
        Private onOption1 As Long
        Private onOption2 As Long
        Private onOption3 As Long
        Private onRemSysType As Long
        Private onBuildingSize As String
        Private ostrCreatedBy As String
        Private ostrDescription As String
        Private ostrManufacturer As String
        Private ostrModifiedBy As String
        Private ostrNotes As String
        Private ostrOwner As String
        Private nBuildingSize As String
        Private strCreatedBy As String
        Private strDescription As String
        Private strManufacturer As String
        Private strModifiedBy As String
        Private strNotes As String
        Private strOwner As String
        Private dtModifiedOn As Date
        'P1 03/09/06 next line only
        Private nOWSSize As String
        Private strOWSManName As String
        Private strOWSSerialNumber As String
        Private strOWSModelNumber As String
        Private nOWSNewUsed As Integer
        'P1 03/09/06 start
        Private nOWSAgeofComp As String
        Private onOWSSize As String
        'P1 03/09/06 end
        Private ostrOWSManName As String
        Private ostrOWSSerialNumber As String
        Private ostrOWSModelNumber As String
        Private onOWSNewUsed As Integer
        'P1 03/09/06 start
        Private onOWSAgeofComp As String

        Private nMotorSize As String
        'P1 03/09/06 end
        Private strMotorManName As String
        Private strMotorSerialNumber As String
        Private strMotorModelNumber As String
        Private nMotorNewUsed As Integer
        'P1 03/09/06 start
        Private nMotorAgeofComp As String
        Private onMotorSize As String
        'P1 03/09/06 end
        Private ostrMotorManName As String
        Private ostrMotorSerialNumber As String
        Private ostrMotorModelNumber As String
        Private onMotorNewUsed As Integer
        'P1 03/09/06 start
        Private onMotorAgeofComp As String

        Private nStripperSize As String
        'P1 03/09/06 end
        Private strStripperManName As String
        Private strStripperSerialNumber As String
        Private strStripperModelNumber As String
        Private nStripperNewUsed As Integer
        'P1 03/09/06 start
        Private nStripperAgeofComp As String
        Private onStripperSize As String
        'P1 03/09/06 end
        Private ostrStripperManName As String
        Private ostrStripperSerialNumber As String
        Private ostrStripperModelNumber As String
        Private onStripperNewUsed As Integer
        'P1 03/09/06 start
        Private onStripperAgeofComp As String

        Private nVacPump1Size As String
        'P1 03/09/06 end
        Private strVacPump1ManName As String
        Private strVacPump1SerialNumber As String
        Private strVacPump1ModelNumber As String
        Private nVacPump1NewUsed As Integer
        Private nVacPump1AgeofComp As String
        Private nVacPump1Seal As Integer
        Private onVacPump1Size As String
        Private ostrVacPump1ManName As String
        Private ostrVacPump1SerialNumber As String
        Private ostrVacPump1ModelNumber As String
        Private onVacPump1NewUsed As Integer
        'P1 03/09/06 next line only
        Private onVacPump1AgeofComp As String
        Private onVacPump1Seal As String
        'P1 03/09/06 next line only
        Private nVacPump2Size As String
        Private strVacPump2ManName As String
        Private strVacPump2SerialNumber As String
        Private strVacPump2ModelNumber As String
        Private nVacPump2NewUsed As Integer
        'P1 03/09/06 next line only
        Private nVacPump2AgeofComp As String
        Private nVacPump2Seal As Integer
        'P1 03/09/06 next line only
        Private onVacPump2Size As String
        Private ostrVacPump2ManName As String
        Private ostrVacPump2SerialNumber As String
        Private ostrVacPump2ModelNumber As String
        Private onVacPump2NewUsed As Integer
        'P1 03/09/06 next line only
        Private onVacPump2AgeofComp As String
        Private onVacPump2Seal As Integer
        Private onEntityID As Integer
        Private nSystemID As Long
        Private nSystemSequence As Long
        Private nSystemDeclaration As Long

        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"

        Sub New()
            MyBase.New()
            dtDataAge = Now()
        End Sub
        'P1 03/09/06 start
        Sub New(ByVal SystemID As Int64, _
                ByVal SystemSequence As Int64, _
                ByVal SystemDeclaration As Int64, _
                ByVal StartDate As Date, _
                ByVal SystemType As Int64, _
                ByVal Description As String, _
                ByVal Manufacturer As String, _
                ByVal OWSSize As String, _
                ByVal OWSManName As String, _
                ByVal OWSSerialNumber As String, _
                ByVal OWSModelNumber As String, _
                ByVal OWSNewUsed As Integer, _
                ByVal OWSCompAge As String, _
                ByVal MotorSize As String, _
                ByVal MotorManName As String, _
                ByVal MotorSerialNumber As String, _
                ByVal MotorModelNumber As String, _
                ByVal MotorNewUsed As Integer, _
                ByVal MotorCompAge As String, _
                ByVal StripperSize As String, _
                ByVal StripperManName As String, _
                ByVal StripperSerialNumber As String, _
                ByVal StripperModelNumber As String, _
                ByVal StripperNewUsed As Integer, _
                ByVal StripperCompAge As String, _
                ByVal VacPump1Size As String, _
                ByVal VacPump1ManName As String, _
                ByVal VacPump1SerialNumber As String, _
                ByVal VacPump1ModelNumber As String, _
                ByVal VacPump1NewUsed As Integer, _
                ByVal VacPump1CompAge As String, _
                ByVal VacPump1Seal As Integer, _
                ByVal VacPump2Size As String, _
                ByVal VacPump2ManName As String, _
                ByVal VacPump2SerialNumber As String, _
                ByVal VacPump2ModelNumber As String, _
                ByVal VacPump2NewUsed As Integer, _
                ByVal VacPump2CompAge As String, _
                ByVal VacPump2Seal As Integer, _
                ByVal Owned As Integer, _
                ByVal SystemOwner As String, _
                ByVal BuildingSize As String, _
                ByVal MountType As Integer, _
                ByVal RefurbishedDate As Date, _
                ByVal Other As String, _
                ByVal Purchase_Date As Date, _
                ByVal CREATED_BY As String, _
                ByVal CREATE_DATE As Date, _
                ByVal LAST_EDITED_BY As String, _
                ByVal DATE_LAST_EDITED As Date, _
                ByVal deleted As Boolean, _
                ByVal Option1 As Long, _
                ByVal Option2 As Long, _
                ByVal Option3 As Long)

            Try
                nSystemID = SystemID
                nSystemSequence = SystemSequence
                nSystemDeclaration = SystemDeclaration
                odtDatePlacedInUse = StartDate.Date
                onRemSysType = SystemType
                ostrDescription = Description
                ostrManufacturer = Manufacturer
                onOWSSize = OWSSize
                ostrOWSManName = OWSManName
                ostrOWSSerialNumber = OWSSerialNumber
                ostrOWSModelNumber = OWSModelNumber
                onOWSNewUsed = OWSNewUsed
                onOWSAgeofComp = OWSCompAge
                onMotorSize = MotorSize
                ostrMotorManName = MotorManName
                ostrMotorSerialNumber = MotorSerialNumber
                ostrMotorModelNumber = MotorModelNumber
                onMotorNewUsed = MotorNewUsed
                onMotorAgeofComp = MotorCompAge
                onStripperSize = StripperSize
                ostrStripperManName = StripperManName
                ostrStripperSerialNumber = StripperSerialNumber
                ostrStripperModelNumber = StripperModelNumber
                onStripperNewUsed = StripperNewUsed
                onStripperAgeofComp = StripperCompAge
                onVacPump1Size = VacPump1Size
                ostrVacPump1ManName = VacPump1ManName
                ostrVacPump1SerialNumber = VacPump1SerialNumber
                ostrVacPump1ModelNumber = VacPump1ModelNumber
                onVacPump1NewUsed = VacPump1NewUsed
                onVacPump1AgeofComp = VacPump1CompAge
                onVacPump1Seal = VacPump1Seal
                onVacPump2Size = VacPump2Size
                ostrVacPump2ManName = VacPump2ManName
                ostrVacPump2SerialNumber = VacPump2SerialNumber
                ostrVacPump2ModelNumber = VacPump2ModelNumber
                onVacPump2NewUsed = VacPump2NewUsed
                onVacPump2AgeofComp = VacPump2CompAge
                onVacPump2Seal = VacPump2Seal
                onOwned = Owned
                ostrOwner = SystemOwner
                onBuildingSize = BuildingSize
                onMountType = MountType
                odtRefurbDate = RefurbishedDate
                ostrNotes = Other
                odtPurchaseDate = Purchase_Date
                ostrCreatedBy = CREATED_BY
                odtCreatedOn = CREATE_DATE
                ostrModifiedBy = LAST_EDITED_BY
                odtModifiedOn = DATE_LAST_EDITED
                obolDeleted = deleted
                onOption1 = Option1
                onOption2 = Option2
                onOption3 = Option3
                dtDataAge = Now()
                Me.Reset()
                'P1 03/09/06 end
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub


#End Region
#Region "Exposed Attributes"
        ' The maximum age the info object can attain before requiring a refresh
        Public Property AgeThreshold() As Integer
            Get
                Return nAgeThreshold
            End Get
            Set(ByVal Value As Integer)
                nAgeThreshold = Value
            End Set
        End Property

        Public Property BuildingSize() As String
            Get
                Return nBuildingSize
            End Get
            Set(ByVal Value As String)
                nBuildingSize = Value
                Me.CheckDirty()
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
        ' The date this remediation system was placed in use.
        Public Property DatePlacedInUse() As Date
            Get
                Return dtDatePlacedInUse
            End Get
            Set(ByVal Value As Date)
                dtDatePlacedInUse = Value
                Me.CheckDirty()
            End Set
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        ' A text description of the remediation sytsem.
        Public Property Description() As String
            Get
                Return strDescription
            End Get
            Set(ByVal Value As String)
                strDescription = Value
                Me.CheckDirty()
            End Set
        End Property
        'ostrManufacturer
        Public Property Manufacturer() As String
            Get
                Return strManufacturer
            End Get
            Set(ByVal Value As String)
                strManufacturer = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The entity ID associated with a technical remediation system.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return nEntityID
            End Get
        End Property
        ' The system ID for this LUST Remediation System
        Public Property ID() As Long
            Get
                Return nSystemID
            End Get
            Set(ByVal Value As Long)
                nSystemID = Value
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

        ' Any additional information required to be known concerning the remediation system.
        Public Property Notes() As String
            Get
                Return strNotes
            End Get
            Set(ByVal Value As String)
                strNotes = Value
                Me.CheckDirty()
            End Set
        End Property
        ' the first piece of optional equipment associated with the remediation system - derived from properties table, property type = "RemSys Option"
        Public Property Option1() As Long
            Get
                Return nOption1
            End Get
            Set(ByVal Value As Long)
                nOption1 = Value
                Me.CheckDirty()
            End Set
        End Property
        ' the second piece of optional equipment associated with the remediation system - derived from properties table, property type = "RemSys Option"
        Public Property Option2() As Long
            Get
                Return nOption2
            End Get
            Set(ByVal Value As Long)
                nOption2 = Value
                Me.CheckDirty()
            End Set
        End Property
        ' the third piece of optional equipment associated with the remediation system - derived from properties table, property type = "RemSys Option"
        Public Property Option3() As Long
            Get
                Return nOption3
            End Get
            Set(ByVal Value As Long)
                nOption3 = Value
                Me.CheckDirty()
            End Set
        End Property
        ' Boolean representing the owned/leased status of the remediation system.
        Public Property Owned() As Integer
            Get
                Return nOwned
            End Get
            Set(ByVal Value As Integer)
                nOwned = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The owner of the remediation system.
        Public Property Owner() As String
            Get
                Return strOwner
            End Get
            Set(ByVal Value As String)
                strOwner = Value
                Me.CheckDirty()
            End Set
        End Property

        ' The purchase date of the remediation system (set to 01/01/01 for leased systems).
        Public Property PurchaseDate() As Date
            Get
                Return dtPurchaseDate
            End Get
            Set(ByVal Value As Date)
                dtPurchaseDate = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The date the remediation system was last refurbished
        Public Property RefurbDate() As Date
            Get
                Return dtRefurbDate
            End Get
            Set(ByVal Value As Date)
                dtRefurbDate = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The type of remediation system - derived from properties table.
        Public Property RemSysType() As Long
            Get
                Return nRemSysType
            End Get
            Set(ByVal Value As Long)
                nRemSysType = Value
                Me.CheckDirty()
            End Set
        End Property

        ' Boolean representing if the system is trailer mounted (false = skid mounted)
        Public Property MountType() As Integer
            Get
                Return nMountType
            End Get
            Set(ByVal Value As Integer)
                nMountType = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OWSSize() As String
            Get
                Return nOWSSize
            End Get
            Set(ByVal Value As String)
                nOWSSize = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OWSManName() As String
            Get
                Return strOWSManName
            End Get
            Set(ByVal Value As String)
                strOWSManName = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OWSSerialNumber() As String
            Get
                Return strOWSSerialNumber
            End Get
            Set(ByVal Value As String)
                strOWSSerialNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OWSModelNumber() As String
            Get
                Return strOWSModelNumber
            End Get
            Set(ByVal Value As String)
                strOWSModelNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OWSNewUsed() As Integer
            Get
                Return nOWSNewUsed
            End Get
            Set(ByVal Value As Integer)
                nOWSNewUsed = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OWSAgeofComp() As String
            Get
                Return nOWSAgeofComp
            End Get
            Set(ByVal Value As String)
                nOWSAgeofComp = Value
                Me.CheckDirty()
            End Set
        End Property


        Public Property MotorSize() As String
            Get
                Return nMotorSize
            End Get
            Set(ByVal Value As String)
                nMotorSize = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property MotorManName() As String
            Get
                Return strMotorManName
            End Get
            Set(ByVal Value As String)
                strMotorManName = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property MotorSerialNumber() As String
            Get
                Return strMotorSerialNumber
            End Get
            Set(ByVal Value As String)
                strMotorSerialNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property MotorModelNumber() As String
            Get
                Return strMotorModelNumber
            End Get
            Set(ByVal Value As String)
                strMotorModelNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property MotorNewUsed() As Integer
            Get
                Return nMotorNewUsed
            End Get
            Set(ByVal Value As Integer)
                nMotorNewUsed = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property MotorAgeofComp() As String
            Get
                Return nMotorAgeofComp
            End Get
            Set(ByVal Value As String)
                nMotorAgeofComp = Value
                Me.CheckDirty()
            End Set
        End Property


        Public Property StripperSize() As String
            Get
                Return nStripperSize
            End Get
            Set(ByVal Value As String)
                nStripperSize = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property StripperManName() As String
            Get
                Return strStripperManName
            End Get
            Set(ByVal Value As String)
                strStripperManName = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property StripperSerialNumber() As String
            Get
                Return strStripperSerialNumber
            End Get
            Set(ByVal Value As String)
                strStripperSerialNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property StripperModelNumber() As String
            Get
                Return strStripperModelNumber
            End Get
            Set(ByVal Value As String)
                strStripperModelNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property StripperNewUsed() As Integer
            Get
                Return nStripperNewUsed
            End Get
            Set(ByVal Value As Integer)
                nStripperNewUsed = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property StripperAgeofComp() As String
            Get
                Return nStripperAgeofComp
            End Get
            Set(ByVal Value As String)
                nStripperAgeofComp = Value
                Me.CheckDirty()
            End Set
        End Property


        Public Property VacPump1Size() As String
            Get
                Return nVacPump1Size
            End Get
            Set(ByVal Value As String)
                nVacPump1Size = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump1ManName() As String
            Get
                Return strVacPump1ManName
            End Get
            Set(ByVal Value As String)
                strVacPump1ManName = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump1SerialNumber() As String
            Get
                Return strVacPump1SerialNumber
            End Get
            Set(ByVal Value As String)
                strVacPump1SerialNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump1ModelNumber() As String
            Get
                Return strVacPump1ModelNumber
            End Get
            Set(ByVal Value As String)
                strVacPump1ModelNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump1NewUsed() As Integer
            Get
                Return nVacPump1NewUsed
            End Get
            Set(ByVal Value As Integer)
                nVacPump1NewUsed = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump1AgeofComp() As String
            Get
                Return nVacPump1AgeofComp
            End Get
            Set(ByVal Value As String)
                nVacPump1AgeofComp = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump1Seal() As Integer
            Get
                Return nVacPump1Seal
            End Get
            Set(ByVal Value As Integer)
                nVacPump1Seal = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump2Size() As String
            Get
                Return nVacPump2Size
            End Get
            Set(ByVal Value As String)
                nVacPump2Size = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump2ManName() As String
            Get
                Return strVacPump2ManName
            End Get
            Set(ByVal Value As String)
                strVacPump2ManName = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump2SerialNumber() As String
            Get
                Return strVacPump2SerialNumber
            End Get
            Set(ByVal Value As String)
                strVacPump2SerialNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump2ModelNumber() As String
            Get
                Return strVacPump2ModelNumber
            End Get
            Set(ByVal Value As String)
                strVacPump2ModelNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump2NewUsed() As Integer
            Get
                Return nVacPump2NewUsed
            End Get
            Set(ByVal Value As Integer)
                nVacPump2NewUsed = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property VacPump2AgeofComp() As String
            Get
                Return nVacPump2AgeofComp
            End Get
            Set(ByVal Value As String)
                nVacPump2AgeofComp = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property VacPump2Seal() As Integer
            Get
                Return nVacPump2Seal
            End Get
            Set(ByVal Value As Integer)
                nVacPump2Seal = Value
                Me.CheckDirty()
            End Set
        End Property


        Public Property SystemSequence() As Long
            Get
                Return nSystemSequence
            End Get
            Set(ByVal Value As Long)
                nSystemSequence = Value
            End Set
        End Property

        Public Property SystemDeclaration() As Long
            Get
                Return nSystemDeclaration
            End Get
            Set(ByVal Value As Long)
                nSystemDeclaration = Value
            End Set
        End Property
#End Region
#Region "Protected Attributes"
        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()

            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = ((odtDatePlacedInUse <> dtDatePlacedInUse) Or _
            (ostrDescription <> strDescription) Or _
            (ostrManufacturer <> strManufacturer) Or _
            (onRemSysType <> nRemSysType) Or _
            (onOWSSize <> nOWSSize) Or _
            (ostrOWSManName <> strOWSManName) Or _
            (ostrOWSSerialNumber <> strOWSSerialNumber) Or _
            (ostrOWSModelNumber <> strOWSModelNumber) Or _
            (onOWSNewUsed <> nOWSNewUsed) Or _
            (onOWSAgeofComp <> nOWSAgeofComp) Or _
            (onMotorSize <> nMotorSize) Or _
            (ostrMotorManName <> strMotorManName) Or _
            (ostrMotorSerialNumber <> strMotorSerialNumber) Or _
            (ostrMotorModelNumber <> strMotorModelNumber) Or _
            (onMotorNewUsed <> nMotorNewUsed) Or _
            (onMotorAgeofComp <> nMotorAgeofComp) Or _
            (onStripperSize <> nStripperSize) Or _
            (ostrStripperManName <> strStripperManName) Or _
            (ostrStripperSerialNumber <> strStripperSerialNumber) Or _
            (ostrStripperModelNumber <> strStripperModelNumber) Or _
            (onStripperNewUsed <> nStripperNewUsed) Or _
            (onStripperAgeofComp <> nStripperAgeofComp) Or _
            (onVacPump1Size <> nVacPump1Size) Or _
            (ostrVacPump1ManName <> strVacPump1ManName) Or _
            (ostrVacPump1SerialNumber <> strVacPump1SerialNumber) Or _
            (ostrVacPump1ModelNumber <> strVacPump1ModelNumber) Or _
            (onVacPump1NewUsed <> nVacPump1NewUsed) Or _
            (onVacPump1AgeofComp <> nVacPump1AgeofComp) Or _
            (onVacPump1Seal <> nVacPump1Seal) Or _
            (onVacPump2Size <> nVacPump2Size) Or _
            (ostrVacPump2ManName <> strVacPump2ManName) Or _
            (ostrVacPump2SerialNumber <> strVacPump2SerialNumber) Or _
            (ostrVacPump2ModelNumber <> strVacPump2ModelNumber) Or _
            (onVacPump2NewUsed <> nVacPump2NewUsed) Or _
            (onVacPump2AgeofComp <> nVacPump2AgeofComp) Or _
            (onVacPump2Seal <> nVacPump2Seal) Or _
            (onOwned <> nOwned) Or _
            (ostrOwner <> strOwner) Or _
            (onBuildingSize <> nBuildingSize) Or _
            (onMountType <> nMountType) Or _
            (odtRefurbDate <> dtRefurbDate) Or _
            (ostrNotes <> strNotes) Or _
            (odtPurchaseDate <> dtPurchaseDate) Or _
            (ostrCreatedBy <> strCreatedBy) Or _
            (odtCreatedOn <> dtCreatedOn) Or _
            (ostrModifiedBy <> strModifiedBy) Or _
            (odtModifiedOn <> dtModifiedOn) Or _
            (obolDeleted <> bolDeleted) Or _
            (onOption1 <> nOption1) Or _
            (onOption2 <> nOption2) Or _
            (onOption3 <> nOption3))


            If bolOldState <> bolIsDirty Then
                RaiseEvent LustRemediationInfoChanged()
            End If
        End Sub


        Private Sub Init()
            nSystemID = 0
            nSystemSequence = 0
            nSystemDeclaration = 0
            onRemSysType = 0
            onEntityID = 0
            odtDatePlacedInUse = "01/01/0001"
            ostrDescription = String.Empty
            ostrManufacturer = String.Empty
            onOWSSize = String.Empty
            ostrOWSManName = String.Empty
            ostrOWSSerialNumber = String.Empty
            ostrOWSModelNumber = String.Empty
            onOWSNewUsed = 0
            onOWSAgeofComp = String.Empty
            onMotorSize = String.Empty
            ostrMotorManName = String.Empty
            ostrMotorSerialNumber = String.Empty
            ostrMotorModelNumber = String.Empty
            onMotorNewUsed = 0
            onMotorAgeofComp = String.Empty
            onStripperSize = String.Empty
            ostrStripperManName = String.Empty
            ostrStripperSerialNumber = String.Empty
            ostrStripperModelNumber = String.Empty
            onStripperNewUsed = 0
            onStripperAgeofComp = String.Empty
            onVacPump1Size = String.Empty
            ostrVacPump1ManName = String.Empty
            ostrVacPump1SerialNumber = String.Empty
            ostrVacPump1ModelNumber = String.Empty
            onVacPump1NewUsed = 0
            onVacPump1AgeofComp = String.Empty
            onVacPump1Seal = 0
            onVacPump2Size = String.Empty
            ostrVacPump2ManName = String.Empty
            ostrVacPump2SerialNumber = String.Empty
            ostrVacPump2ModelNumber = String.Empty
            onVacPump2NewUsed = 0
            onVacPump2AgeofComp = String.Empty
            onVacPump2Seal = 0
            onOwned = 0
            ostrOwner = String.Empty
            onBuildingSize = String.Empty
            onMountType = 0
            odtRefurbDate = "01/01/0001"
            ostrNotes = String.Empty
            odtPurchaseDate = "01/01/0001"
            ostrCreatedBy = String.Empty
            odtCreatedOn = "01/01/0001"
            ostrModifiedBy = String.Empty
            odtModifiedOn = "01/01/0001"
            obolDeleted = False
            onOption1 = 0
            onOption2 = 0
            onOption3 = 0


            Me.Reset()
        End Sub

        Public Sub Reset()

            Try
                dtDatePlacedInUse = odtDatePlacedInUse
                nRemSysType = onRemSysType
                nEntityID = onEntityID
                strDescription = ostrDescription
                strManufacturer = ostrManufacturer
                nOWSSize = onOWSSize
                strOWSManName = ostrOWSManName
                strOWSSerialNumber = ostrOWSSerialNumber
                strOWSModelNumber = ostrOWSModelNumber
                nOWSNewUsed = onOWSNewUsed
                nOWSAgeofComp = onOWSAgeofComp
                nMotorSize = onMotorSize
                strMotorManName = ostrMotorManName
                strMotorSerialNumber = ostrMotorSerialNumber
                strMotorModelNumber = ostrMotorModelNumber
                nMotorNewUsed = onMotorNewUsed
                nMotorAgeofComp = onMotorAgeofComp
                nStripperSize = onStripperSize
                strStripperManName = ostrStripperManName
                strStripperSerialNumber = ostrStripperSerialNumber
                strStripperModelNumber = ostrStripperModelNumber
                nStripperNewUsed = onStripperNewUsed
                nStripperAgeofComp = onStripperAgeofComp
                nVacPump1Size = onVacPump1Size
                strVacPump1ManName = ostrVacPump1ManName
                strVacPump1SerialNumber = ostrVacPump1SerialNumber
                strVacPump1ModelNumber = ostrVacPump1ModelNumber
                nVacPump1NewUsed = onVacPump1NewUsed
                nVacPump1AgeofComp = onVacPump1AgeofComp
                nVacPump1Seal = onVacPump1Seal
                nVacPump2Size = onVacPump2Size
                strVacPump2ManName = ostrVacPump2ManName
                strVacPump2SerialNumber = ostrVacPump2SerialNumber
                strVacPump2ModelNumber = ostrVacPump2ModelNumber
                nVacPump2NewUsed = onVacPump2NewUsed
                nVacPump2AgeofComp = onVacPump2AgeofComp
                nVacPump2Seal = onVacPump2Seal
                nOwned = onOwned
                strOwner = ostrOwner
                nBuildingSize = onBuildingSize
                nMountType = onMountType
                dtRefurbDate = odtRefurbDate
                strNotes = ostrNotes
                dtPurchaseDate = odtPurchaseDate
                strCreatedBy = ostrCreatedBy
                dtCreatedOn = odtCreatedOn
                strModifiedBy = ostrModifiedBy
                dtModifiedOn = odtModifiedOn
                bolDeleted = obolDeleted
                nOption1 = onOption1
                nOption2 = onOption2
                nOption3 = onOption3


            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try


        End Sub
        Public Sub Archive()

            Try
                odtDatePlacedInUse = dtDatePlacedInUse
                onEntityID = nEntityID
                onRemSysType = nRemSysType
                ostrDescription = strDescription
                ostrManufacturer = strManufacturer
                onOWSSize = nOWSSize
                ostrOWSManName = strOWSManName
                ostrOWSSerialNumber = strOWSSerialNumber
                ostrOWSModelNumber = strOWSModelNumber
                onOWSNewUsed = nOWSNewUsed
                onOWSAgeofComp = nOWSAgeofComp
                onMotorSize = nMotorSize
                ostrMotorManName = strMotorManName
                ostrMotorSerialNumber = strMotorSerialNumber
                ostrMotorModelNumber = strMotorModelNumber
                onMotorNewUsed = nMotorNewUsed
                onMotorAgeofComp = nMotorAgeofComp
                onStripperSize = nStripperSize
                ostrStripperManName = strStripperManName
                ostrStripperSerialNumber = strStripperSerialNumber
                ostrStripperModelNumber = strStripperModelNumber
                onStripperNewUsed = nStripperNewUsed
                onStripperAgeofComp = nStripperAgeofComp
                onVacPump1Size = nVacPump1Size
                ostrVacPump1ManName = strVacPump1ManName
                ostrVacPump1SerialNumber = strVacPump1SerialNumber
                ostrVacPump1ModelNumber = strVacPump1ModelNumber
                onVacPump1NewUsed = nVacPump1NewUsed
                onVacPump1AgeofComp = nVacPump1AgeofComp
                onVacPump1Seal = nVacPump1Seal
                onVacPump2Size = nVacPump2Size
                ostrVacPump2ManName = strVacPump2ManName
                ostrVacPump2SerialNumber = strVacPump2SerialNumber
                ostrVacPump2ModelNumber = strVacPump2ModelNumber
                onVacPump2NewUsed = nVacPump2NewUsed
                onVacPump2AgeofComp = nVacPump2AgeofComp
                onVacPump2Seal = nVacPump2Seal
                onOwned = nOwned
                ostrOwner = strOwner
                onBuildingSize = nBuildingSize
                onMountType = nMountType
                odtRefurbDate = dtRefurbDate
                ostrNotes = strNotes
                odtPurchaseDate = dtPurchaseDate
                ostrCreatedBy = strCreatedBy
                odtCreatedOn = dtCreatedOn
                ostrModifiedBy = strModifiedBy
                odtModifiedOn = dtModifiedOn
                obolDeleted = bolDeleted
                onOption1 = nOption1
                onOption2 = nOption2
                onOption3 = nOption3

            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

    End Class
End Namespace
