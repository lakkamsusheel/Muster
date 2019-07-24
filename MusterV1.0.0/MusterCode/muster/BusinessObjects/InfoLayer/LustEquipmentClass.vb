
Namespace MUSTER.Info
    Public Class LustEquipmentClass
        Public Delegate Sub InfoChangedEventHandler()
        Public Sub New()
        End Sub
        ' The maximum age the info object can attain before requiring a refresh
        Public Property AgeThreshold() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{4A717F21-3341-4208-A483-9A8BACA51984}
                Return dtDataAge
                ' #End Region ' XDEOperation End Template Expansion{4A717F21-3341-4208-A483-9A8BACA51984}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{81C833E2-220F-421E-9A7A-3B2CC3B9FB39}
                dtDataAge = Value
                ' #End Region ' XDEOperation End Template Expansion{81C833E2-220F-421E-9A7A-3B2CC3B9FB39}
            End Set
        End Property
        Public Sub Archive()
        End Sub
        ' The age of the components of the unit
        Public Property ComponentsAge() As Double
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{8E50B506-44CF-430B-9825-3E04C9D1A7A7}
                Return nComponentsAge
                ' #End Region ' XDEOperation End Template Expansion{8E50B506-44CF-430B-9825-3E04C9D1A7A7}
            End Get
            Set(ByVal Value As Double)
                ' #Region "XDEOperation" ' Begin Template Expansion{4505B681-FBED-40F4-9F6F-36F0564C032F}
                nComponentsAge = Value
                ' #End Region ' XDEOperation End Template Expansion{4505B681-FBED-40F4-9F6F-36F0564C032F}
            End Set
        End Property
        ' The ID of the user that created the row
        Public ReadOnly Property CreatedBy() As String
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{3C051DD3-2D74-4F71-9C3C-30F06ACC7E78}
                Return strCreatedBy
                ' #End Region ' XDEOperation End Template Expansion{3C051DD3-2D74-4F71-9C3C-30F06ACC7E78}
            End Get
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{F0EC3ABE-53C5-4ECB-9ED0-7AACB21427B7}
                Return dtCreatedOn
                ' #End Region ' XDEOperation End Template Expansion{F0EC3ABE-53C5-4ECB-9ED0-7AACB21427B7}
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{62C7C52B-FE4E-49F3-AB97-7A51ADFAB009}
                Return bolDeleted
                ' #End Region ' XDEOperation End Template Expansion{62C7C52B-FE4E-49F3-AB97-7A51ADFAB009}
            End Get
            Set(ByVal Value As Boolean)
                ' #Region "XDEOperation" ' Begin Template Expansion{F45D08F9-6F85-4F50-8CC1-7F38CED752CF}
                bolDeleted = Value
                ' #End Region ' XDEOperation End Template Expansion{F45D08F9-6F85-4F50-8CC1-7F38CED752CF}
            End Set
        End Property
        ' The entity ID associated with technical equipment.
        Public ReadOnly Property EntityID() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{960B6F22-0703-4492-B891-8DB6468B000B}
                Return nEntityID
                ' #End Region ' XDEOperation End Template Expansion{960B6F22-0703-4492-B891-8DB6468B000B}
            End Get
        End Property
        ' The system ID for this LUST equipment.
        Public ReadOnly Property ID() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{A0B81136-12BB-4719-8E07-49B06333F62F}

                ' #End Region ' XDEOperation End Template Expansion{A0B81136-12BB-4719-8E07-49B06333F62F}
            End Get
        End Property
        ' Raised when any of the LustEventInfo attributes are modified
        Public Event InfoChanged As InfoChangedEventHandler
        Public Property IsDirty() As Boolean
            Get
            End Get
            Set(ByVal Value As Boolean)
            End Set
        End Property
        ' the manufacturer of the unit.
        Public Property Manufacturer() As String
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{37DE3219-F156-4F65-B1DB-FEFE5ABBF88E}
                Return strManufacturer
                ' #End Region ' XDEOperation End Template Expansion{37DE3219-F156-4F65-B1DB-FEFE5ABBF88E}
            End Get
            Set(ByVal Value As String)
                ' #Region "XDEOperation" ' Begin Template Expansion{6FEEAC3C-ADBF-49C0-B86F-C0AD0CE44803}
                strManufacturer = Value
                ' #End Region ' XDEOperation End Template Expansion{6FEEAC3C-ADBF-49C0-B86F-C0AD0CE44803}
            End Set
        End Property
        ' The model number of the unit
        Public Property ModelNumber() As String
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{02B453BC-BC3F-4C1C-A7E9-26D0B3F35B52}
                Return strModelNumber
                ' #End Region ' XDEOperation End Template Expansion{02B453BC-BC3F-4C1C-A7E9-26D0B3F35B52}
            End Get
            Set(ByVal Value As String)
                ' #Region "XDEOperation" ' Begin Template Expansion{A261C348-48EE-4E03-8E01-F81DE7582F37}
                strModelNumber = Value
                ' #End Region ' XDEOperation End Template Expansion{A261C348-48EE-4E03-8E01-F81DE7582F37}
            End Set
        End Property
        Public ReadOnly Property ModifiedBy() As String
            Get
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
            End Get
        End Property
        ' Boolean indicating the used/new status of the unit.
        Public Property NewUnit() As Boolean
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{B0728BBF-508C-45CB-B564-AA7DF74960E3}
                Return bolNewUnit
                ' #End Region ' XDEOperation End Template Expansion{B0728BBF-508C-45CB-B564-AA7DF74960E3}
            End Get
            Set(ByVal Value As Boolean)
                ' #Region "XDEOperation" ' Begin Template Expansion{39197CBE-B658-41BF-8027-5883C42FD08B}
                bolNewUnit = Value
                ' #End Region ' XDEOperation End Template Expansion{39197CBE-B658-41BF-8027-5883C42FD08B}
            End Set
        End Property
        Public Sub Reset()
        End Sub
        ' The seal type for the unit's pump (oil/water) - derived from properties table, property type "Pump Seal"
        Public Property SealType() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{A7580A4A-C998-4D4B-888C-DDCC42FE91E2}
                Return nSealType
                ' #End Region ' XDEOperation End Template Expansion{A7580A4A-C998-4D4B-888C-DDCC42FE91E2}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{64F75540-4B09-4184-ADF8-CF3612E94A0B}
                nSealType = Value
                ' #End Region ' XDEOperation End Template Expansion{64F75540-4B09-4184-ADF8-CF3612E94A0B}
            End Set
        End Property
        ' The serial number of the unit.
        Public Property SerialNumber() As String
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{D96DC6FF-8FB2-41EE-A7E6-4BEC7F2206B5}
                Return strSerialNumber
                ' #End Region ' XDEOperation End Template Expansion{D96DC6FF-8FB2-41EE-A7E6-4BEC7F2206B5}
            End Get
            Set(ByVal Value As String)
                ' #Region "XDEOperation" ' Begin Template Expansion{F6F28BAD-88C0-477B-BB9E-AA10FD1E2F16}
                strSerialNumber = Value
                ' #End Region ' XDEOperation End Template Expansion{F6F28BAD-88C0-477B-BB9E-AA10FD1E2F16}
            End Set
        End Property
        ' The size of the unit.
        Public Property UnitSize() As Double
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{DAE73247-9353-4516-ADA4-7039FF57FA0D}
                Return nUnitSize
                ' #End Region ' XDEOperation End Template Expansion{DAE73247-9353-4516-ADA4-7039FF57FA0D}
            End Get
            Set(ByVal Value As Double)
                ' #Region "XDEOperation" ' Begin Template Expansion{8822483F-3C80-4C42-9053-8C8396D0C4B3}
                nUnitSize = Value
                ' #End Region ' XDEOperation End Template Expansion{8822483F-3C80-4C42-9053-8C8396D0C4B3}
            End Set
        End Property
        Protected Overrides Sub Finalize()
        End Sub
        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{07B7B8F4-0B24-4EA0-B546-4C0D4AB74F19}
                Return True
                ' #End Region ' XDEOperation End Template Expansion{07B7B8F4-0B24-4EA0-B546-4C0D4AB74F19}
            End Get
        End Property
        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean
        Private bolNewUnit As Boolean
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private nComponentsAge As Double
        Private nEntityID As Integer
        Private nSealType As Long
        Private nUnitSize As Double
        Private obolDeleted As Boolean
        Private obolIsDirty As Boolean
        Private obolNewUnit As Boolean
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private onComponentsAge As Double
        Private onSealType As Long
        Private onUnitSize As Double
        Private ostrCreatedBy As String
        Private ostrManufacturer As String
        Private ostrModelNumber As String
        Private ostrModifiedBy As String
        Private ostrSerialNumber As String
        Private strCreatedBy As String
        Private strManufacturer As String
        Private strModelNumber As String
        Private strModifiedBy As String
        Private strSerialNumber As String
        Private Sub CheckDirty()
        End Sub
        Private Sub Init()
        End Sub
    End Class
End Namespace
