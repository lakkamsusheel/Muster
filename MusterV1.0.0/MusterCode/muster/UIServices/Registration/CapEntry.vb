Imports Infragistics.Win
Imports Infragistics.Win.UltraWinGrid
Imports System.Data.SqlClient
Imports System.Text


Public Class CapEntry
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents pnlCenterRestTop As System.Windows.Forms.Panel
    Friend WithEvents btnCopyCurrenttoNext As System.Windows.Forms.Button
    Friend WithEvents btnApplytoAll As System.Windows.Forms.Button
    Friend WithEvents btnExpand As System.Windows.Forms.Button
    Friend WithEvents btnCollapse As System.Windows.Forms.Button
    Friend WithEvents ugTankandPipe As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.pnlCenterRestTop = New System.Windows.Forms.Panel
        Me.btnCopyCurrenttoNext = New System.Windows.Forms.Button
        Me.btnApplytoAll = New System.Windows.Forms.Button
        Me.btnExpand = New System.Windows.Forms.Button
        Me.btnCollapse = New System.Windows.Forms.Button
        Me.ugTankandPipe = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.pnlCenterRestTop.SuspendLayout()
        CType(Me.ugTankandPipe, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlCenterRestTop
        '
        Me.pnlCenterRestTop.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlCenterRestTop.Controls.Add(Me.btnCopyCurrenttoNext)
        Me.pnlCenterRestTop.Controls.Add(Me.btnApplytoAll)
        Me.pnlCenterRestTop.Controls.Add(Me.btnExpand)
        Me.pnlCenterRestTop.Controls.Add(Me.btnCollapse)
        Me.pnlCenterRestTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlCenterRestTop.Name = "pnlCenterRestTop"
        Me.pnlCenterRestTop.Size = New System.Drawing.Size(888, 38)
        Me.pnlCenterRestTop.TabIndex = 127
        '
        'btnCopyCurrenttoNext
        '
        Me.btnCopyCurrenttoNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopyCurrenttoNext.Location = New System.Drawing.Point(496, 8)
        Me.btnCopyCurrenttoNext.Name = "btnCopyCurrenttoNext"
        Me.btnCopyCurrenttoNext.Size = New System.Drawing.Size(120, 23)
        Me.btnCopyCurrenttoNext.TabIndex = 120
        Me.btnCopyCurrenttoNext.Text = "Copy Current to Next"
        '
        'btnApplytoAll
        '
        Me.btnApplytoAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApplytoAll.Location = New System.Drawing.Point(624, 8)
        Me.btnApplytoAll.Name = "btnApplytoAll"
        Me.btnApplytoAll.Size = New System.Drawing.Size(120, 23)
        Me.btnApplytoAll.TabIndex = 122
        Me.btnApplytoAll.Text = "Apply To All"
        '
        'btnExpand
        '
        Me.btnExpand.Location = New System.Drawing.Point(8, 8)
        Me.btnExpand.Name = "btnExpand"
        Me.btnExpand.TabIndex = 123
        Me.btnExpand.Text = "Expand All"
        '
        'btnCollapse
        '
        Me.btnCollapse.Location = New System.Drawing.Point(88, 8)
        Me.btnCollapse.Name = "btnCollapse"
        Me.btnCollapse.TabIndex = 124
        Me.btnCollapse.Text = "Collapse All"
        '
        'ugTankandPipe
        '
        Me.ugTankandPipe.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ugTankandPipe.Cursor = System.Windows.Forms.Cursors.Default
        Appearance1.FontData.BoldAsString = "True"
        Me.ugTankandPipe.DisplayLayout.CaptionAppearance = Appearance1
        Me.ugTankandPipe.Location = New System.Drawing.Point(0, 40)
        Me.ugTankandPipe.Name = "ugTankandPipe"
        Me.ugTankandPipe.Size = New System.Drawing.Size(888, 312)
        Me.ugTankandPipe.TabIndex = 126
        Me.ugTankandPipe.Text = "Tank / Pipe"
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BackColor = System.Drawing.Color.Cornsilk
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Location = New System.Drawing.Point(2, 383)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(884, 23)
        Me.Label1.TabIndex = 128
        Me.Label1.Text = "Label1"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(728, 352)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 129
        Me.btnSave.Text = "Save"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(808, 352)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 130
        Me.btnCancel.Text = "Cancel"
        '
        'CapEntry
        '
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pnlCenterRestTop)
        Me.Controls.Add(Me.ugTankandPipe)
        Me.Name = "CapEntry"
        Me.Size = New System.Drawing.Size(888, 408)
        Me.pnlCenterRestTop.ResumeLayout(False)
        CType(Me.ugTankandPipe, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "User Defined Variables"

    Public CleanFieldsOnLoad As Boolean = False
    Public SaveAll As Boolean = False

    Public Event Fillfacilities(ByVal facID As Integer)

    Public CAP_STATUS As Boolean = False
    Private bolSelfOwner As Boolean = False
    Private dctCapFields As Collections.Specialized.ListDictionary

    Dim WithEvents dsCAPTankandPipe As New DataSet
    Event CapChanged()


    Private WithEvents pTank As New MUSTER.BusinessLogic.pTank
    Private WithEvents pPipe As New MUSTER.BusinessLogic.pPipe
    Private oOwner As BusinessLogic.pOwner
    Private isDirty As Boolean = False
    Private bolUseInspectionMode As Boolean = False

    Private bolEvaluatingCellValue As Boolean = False
    Private bolEvaluatingCellValueInProgress As Boolean = False

    Private bolDisplayErrmessage As Boolean
    Private bolValidateSuccess As Boolean

    Friend CallingForm As Form

    Private bolFacCapStatusChanged As Boolean = False
    Dim returnVal As String = String.Empty

    Friend nOwnerID As Integer
    Friend nFacilityID As Integer

    Public ReadOnly Property CapDictionary() As Collections.Specialized.ListDictionary

        Get
            If Me.dctCapFields Is Nothing Then
                Me.dctCapFields = New Collections.Specialized.ListDictionary
            End If

            Return Me.dctCapFields
        End Get

    End Property


    Public ReadOnly Property CheckForDirtyScreen() As Boolean
        Get

            If isDirty Then

                If (MsgBox("You made changes to some CAP fields. Do you wish to save your CAP changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
                    If ApplyChanges() Then
                        Return False
                    Else
                        Return True
                    End If
                End If

                Return False
            Else
                Return False
            End If

        End Get
    End Property
#End Region


    Sub New(ByVal oownerInfo As BusinessLogic.pOwner, ByVal facilityID As Integer, ByVal thisCallingTag As Form)

        Call Me.New()

        oOwner = oownerInfo
        CallingForm = thisCallingTag
        nOwnerID = oOwner.ID
        nFacilityID = facilityID

    End Sub

    Sub New(ByVal oownerID As Integer, ByVal facilityID As Integer, ByVal thisCallingTag As Form, Optional ByVal useInspection As Boolean = False)

        Call Me.New()

        bolUseInspectionMode = useInspection

        oOwner = New BusinessLogic.pOwner
        oOwner.Facilities = New BusinessLogic.pFacility

        bolSelfOwner = True

        If useInspection Then
            btnCopyCurrenttoNext.Visible = False
        End If

        CallingForm = thisCallingTag
        nOwnerID = oownerID
        nFacilityID = facilityID

    End Sub

#Region "Form Page Events "

    Public Sub LoadForm()
        LoadTankandPipe(nFacilityID)
    End Sub

    Private Sub Control_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim strParticipationLevel As String
        Try

            Dim strAddress As String = String.Empty

            LoadForm()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Overloads Sub dispose()
        ugTankandPipe.DataSource = Nothing
        ugTankandPipe.ResetDisplayLayout()
        ugTankandPipe.Layouts.Clear()

        If bolSelfOwner Then
            oOwner = Nothing
        End If

    End Sub

#End Region


#Region " Tank/Pipe Grid Events & Processes "

    Dim cellRed As Boolean

    Public Sub LoadTankandPipe(ByVal nSelectedFacID As Integer)
        Try



            If Not IsDBNull(nSelectedFacID) Then
                Dim fac As New BusinessLogic.pFacility

                dsCAPTankandPipe = fac.getCAPTanksandPipesByFacility(nSelectedFacID, bolUseInspectionMode, Me.CleanFieldsOnLoad)

                If CleanFieldsOnLoad Then
                    CleanFieldsOnLoad = False
                    SaveAll = True
                End If


                fac = Nothing

                If dsCAPTankandPipe.Tables(0).Rows.Count <= 0 Then
                    MsgBox("No Records Found")

                    ugTankandPipe.DataSource = Nothing
                    ugTankandPipe.ResetDisplayLayout()
                    ugTankandPipe.Layouts.Clear()

                    Exit Sub

                Else

                    ugTankandPipe.DataSource = Nothing
                    FillugTankAndPipeGrid(dsCAPTankandPipe)

                End If

                isDirty = False
                UpdateLabel()

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FillugTankAndPipeGrid(ByRef dtSet As DataSet)

        Try
            ' Hiding Tank Columns  
            ugTankandPipe.DataSource = Nothing


            With dsCAPTankandPipe

                'tanks
                If .Tables.Count > 0 Then
                    With .Tables(0)

                        .Columns("FACILITY_ID").ColumnMapping = MappingType.Hidden
                        .Columns("TANK ID").ColumnMapping = MappingType.Hidden
                        .Columns("TANKLD").ColumnMapping = MappingType.Hidden
                        .Columns("TANKMODDESC").ColumnMapping = MappingType.Hidden
                        .Columns("TCPINSTALLDATE").ColumnMapping = MappingType.Hidden
                        .Columns("TANK CP TYPE").ColumnMapping = MappingType.Hidden
                        .Columns("SMALLDELIVERY").ColumnMapping = MappingType.Hidden
                        .Columns("TANKEMERGEN").ColumnMapping = MappingType.Hidden

                    End With
                End If


                If .Tables.Count > 1 Then

                    'Pipes
                    With .Tables(1)

                        .Columns("FACILITY_ID").ColumnMapping = MappingType.Hidden
                        .Columns("TANK ID").ColumnMapping = MappingType.Hidden
                        .Columns("COMPARTMENT_NUMBER").ColumnMapping = MappingType.Hidden
                        .Columns("PIPE ID").ColumnMapping = MappingType.Hidden
                        .Columns("PIPE_MOD_DESC").ColumnMapping = MappingType.Hidden
                        .Columns("PIPE_LD").ColumnMapping = MappingType.Hidden
                        .Columns("PIPE_TYPE_DESC").ColumnMapping = MappingType.Hidden
                        .Columns("TERMINATION_TYPE_DISP").ColumnMapping = MappingType.Hidden
                        .Columns("TERMINATION_TYPE_TANK").ColumnMapping = MappingType.Hidden
                        .Columns("PIPE CP TYPE").ColumnMapping = MappingType.Hidden
                        .Columns("PIPE_CP_INSTALLED_DATE").ColumnMapping = MappingType.Hidden
                        .Columns("TERMINATION_CP_INSTALLED_DATE").ColumnMapping = MappingType.Hidden

                    End With
                End If
            End With

            With ugTankandPipe

                .DataSource = dsCAPTankandPipe
                .Rows.ExpandAll(True)

                With .DisplayLayout

                    .Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False

                    With .Bands(0)

                        .Columns("CP Date").Header.Caption = "Tank CP Tested"

                        .Columns("TANK site Id").SortIndicator = SortIndicator.Ascending


                        .Columns("INSTALLED").Hidden = True
                        .Columns("CAPACITY").Hidden = True
                        .Columns("SUBSTANCE").Hidden = True
                        .Columns("LI INSTALL").Hidden = True

                        '.Columns("DateSecondaryContainmentLastInspected").Header.Caption = "Tank Sec Insp"
                        .Columns("DateSecondaryContainmentLastInspected").Hidden = True
                        .Columns("Tank_LD_Num").Hidden = True

                        .Columns("DateSpillPreventionInstalled").Hidden = True
                        .Columns("DateOverfillPreventionInstalled").Hidden = True
                        .Columns("DateSpillPreventionLastTested").Header.Caption = "Spill Tested"
                        .Columns("DateOverfillPreventionLastInspected").Header.Caption = "Overfill Tested"
                        .Columns("DateElectronicDeviceInspected").Header.Caption = "Tank Elec Inspected"
                        .Columns("LI INSPECTED").Header.Caption = "Lining Inspected"
                        .Columns("TT Date").Header.Caption = "TT Date"
                        .Columns("DateATGLastInspected").Header.Caption = "Auto-Tank Gauging Inspected"




                    End With

                    With .Bands(1)

                        .Columns("DISP CP TYPE").Hidden = True
                        .Columns("CP Date").Header.Caption = "Pipe CP Tested"
                        .Columns("INSTALL DATE").Hidden = True
                        .Columns("ALLD_Test").Hidden = True
                        .Columns("TANK CP TYPE").Hidden = True

                        'ALLD_TEST_DATE
                        .Columns("ALLD_TEST_DATE").Header.Caption = "ALLD Tested"
                        .Columns("ALLD_Test").Header.Caption = "ALLD Type"

                        .Columns("DateSheerValueTest").Header.Caption = "Shear Tested"
                        .Columns("TERM CP TEST").Header.Caption = "Term CP Tested"
                        .Columns("DateSecondaryContainmentInspect").Header.Caption = "Secondary Inspected"
                        .Columns("DateElectronicDeviceInspect").Header.Caption = "Pipe Elec Inspected"
                        .Columns("TT Date").Header.Caption = "TT Date"
                        .Columns("PipeType").Hidden = True
                        .Columns("Pipe_LD_Num").Hidden = True

                    End With
                End With

                Refresh()

            End With

            CapDictionary.Clear()

            For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In Me.ugTankandPipe.Rows

                For Each row2 As Infragistics.Win.UltraWinGrid.UltraGridRow In row.ChildBands(0).Rows

                    Dim PipeID As String = String.Format("P{0}", row2.Cells("PIPE ID").Value)

                    For Each col As Infragistics.Win.UltraWinGrid.UltraGridCell In row2.Cells
                        If CapDictionary.Item(String.Format("{0}_{1}", PipeID, col.Column.Header.Caption).ToUpper) Is Nothing Then
                            CapDictionary.Add(String.Format("{0}_{1}", PipeID, col.Column.Header.Caption).ToUpper, col)
                        Else
                            CapDictionary.Item(String.Format("{0}_{1}", PipeID, col.Column.Header.Caption).ToUpper) = col
                        End If

                    Next
                Next

                Dim tankID As String = String.Format("T{0}", row.Cells("TANK ID").Value)

                For Each col As Infragistics.Win.UltraWinGrid.UltraGridCell In row.Cells
                    If CapDictionary.Item(String.Format("{0}_{1}", tankID, col.Column.Header.Caption).ToUpper) Is Nothing Then
                        CapDictionary.Add(String.Format("{0}_{1}", tankID, col.Column.Header.Caption).ToUpper, col)
                    Else
                        CapDictionary.Item(String.Format("{0}_{1}", tankID, col.Column.Header.Caption).ToUpper) = col
                    End If
                Next

            Next




            ApplyBusinessRules()

            Refresh()

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub btnExpand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpand.Click
        ugTankandPipe.Rows.ExpandAll(True)
    End Sub

    Private Sub btnCollapse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCollapse.Click
        ugTankandPipe.Rows.CollapseAll(True)
    End Sub

    Public Sub ApplyBusinessRules()

        Dim dCol As DataColumn
        Dim dttemp, dtValiddate, dtNullDate As Date
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugchildrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim fac As New BusinessLogic.pFacility
        Dim dtTodayPlus90Days As Date = Today

        cellRed = False

        Try


            'Starts here
            For Each ugrow In ugTankandPipe.Rows ' tank
                If Not (ugrow.Cells("Substance").Text.ToString = "Used Oil") Then
                    If Not (ugrow.Cells("SmallDelivery").Text.ToString = "True") Then
                        If ugrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then

                            If ugrow.Cells("DateSpillPreventionInstalled").Value Is DBNull.Value Then
                                ugrow.Cells("DateSpillPreventionInstalled").Appearance.BackColor = Color.Red
                            Else
                                ugrow.Cells("DateSpillPreventionInstalled").Appearance.BackColor = Color.Yellow
                            End If

                            If ugrow.Cells("DateOverfillPreventionInstalled").Value Is DBNull.Value Then
                                ugrow.Cells("DateOverfillPreventionInstalled").Appearance.BackColor = Color.Red
                            Else
                                ugrow.Cells("DateOverfillPreventionInstalled").Appearance.BackColor = Color.Yellow
                            End If


                            If ugrow.Cells("DateSpillPreventionLastTested").Value Is DBNull.Value Then

                                ugrow.Cells("DateSpillPreventionLastTested").Appearance.BackColor = Color.Red
                                cellRed = True
                            Else

                                dttemp = ugrow.Cells("DateSpillPreventionLastTested").Value
                                dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                                dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                                If Date.Compare(dttemp, Today) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                    ugrow.Cells("DateSpillPreventionLastTested").Appearance.BackColor = Color.Red
                                    cellRed = True
                                Else
                                    ugrow.Cells("DateSpillPreventionLastTested").Appearance.BackColor = Color.Yellow
                                End If
                            End If

                            If ugrow.Cells("DateOverfillPreventionLastInspected").Value Is DBNull.Value Then
                                ugrow.Cells("DateOverfillPreventionLastInspected").Appearance.BackColor = Color.Red
                                cellRed = True

                            Else

                                dttemp = ugrow.Cells("DateOverfillPreventionLastInspected").Value
                                dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                                dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                                If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                    ugrow.Cells("DateOverfillPreventionLastInspected").Appearance.BackColor = Color.Red
                                    cellRed = True
                                Else
                                    ugrow.Cells("DateOverfillPreventionLastInspected").Appearance.BackColor = Color.Yellow
                                End If

                            End If
                        End If

                    Else

                        ugrow.Cells("DateSpillPreventionInstalled").Appearance.BackColor = Color.White

                        ugrow.Cells("DateOverfillPreventionInstalled").Appearance.BackColor = Color.White
                        ugrow.Cells("DateSpillPreventionLastTested").Appearance.BackColor = Color.White
                        ugrow.Cells("DateOverfillPreventionLastInspected").Appearance.BackColor = Color.White

                    End If

                    If (ugrow.Cells("TankLD").Text.ToString = "Electronic Interstitial Monitoring") And ugrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then

                        If ugrow.Cells("DateElectronicDeviceInspected").Value Is DBNull.Value Then

                            ugrow.Cells("DateElectronicDeviceInspected").Appearance.BackColor = Color.Red
                            cellRed = True

                        Else

                            dttemp = ugrow.Cells("DateElectronicDeviceInspected").Value
                            dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                            dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                            dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                            If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                ugrow.Cells("DateElectronicDeviceInspected").Appearance.BackColor = Color.Red
                                cellRed = True
                            Else
                                ugrow.Cells("DateElectronicDeviceInspected").Appearance.BackColor = Color.Yellow
                            End If

                        End If

                    Else
                        ugrow.Cells("DateElectronicDeviceInspected").Appearance.BackColor = Color.White
                    End If

                    If (ugrow.Cells("TankLD").Text.ToString = "Automatic Tank Gauging") And ugrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then

                        If ugrow.Cells("DateATGLastInspected").Value Is DBNull.Value Then
                            ugrow.Cells("DateATGLastInspected").Appearance.BackColor = Color.Red
                            cellRed = True

                        Else

                            dttemp = ugrow.Cells("DateATGLastInspected").Value
                            dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                            dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                            dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)

                            If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                ugrow.Cells("DateATGLastInspected").Appearance.BackColor = Color.Red
                                cellRed = True

                            Else
                                ugrow.Cells("DateATGLastInspected").Appearance.BackColor = Color.Yellow
                            End If

                        End If

                    Else
                        ugrow.Cells("DateATGLastInspected").Appearance.BackColor = Color.White
                    End If

                End If


                'If Tank Status is CIU
                If ugrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then

                    'For Tank
                    If ugrow.Cells("TANKLD").Text.IndexOf("Inventory Control/Precision Tightness Testing") >= 0 Then

                        ugrow.Cells("TT DATE").Appearance.BackColor = Color.Yellow

                        If ugrow.Cells("TT DATE").Value Is DBNull.Value Then
                            ugrow.Cells("TT DATE").Appearance.BackColor = Color.Red

                        Else
                            dttemp = ugrow.Cells("TT DATE").Value
                            dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                            dtValiddate = DateAdd(DateInterval.Year, -5, dtValiddate)
                            dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                            If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                ugrow.Cells("TT DATE").Appearance.BackColor = Color.Red
                            End If
                        End If

                    Else

                        ugrow.Cells("TT DATE").Appearance.BackColor = Color.White
                    End If

                End If


                'If Tank Status is CIU or TOSI
                If ugrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Or ugrow.Cells("STATUS").Text.IndexOf("Temporarily Out of Service Indefinitely") >= 0 Then

                    If ugrow.Cells("TANKMODDESC").Text.IndexOf("Cathodically Protected") >= 0 Then

                        ugrow.Cells("CP DATE").Appearance.BackColor = Color.Yellow

                        If ugrow.Cells("CP DATE").Value Is DBNull.Value Then
                            ugrow.Cells("CP DATE").Appearance.BackColor = Color.Red

                        Else
                            dttemp = ugrow.Cells("CP DATE").Value
                            dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                            dtValiddate = DateAdd(DateInterval.Year, -3, dtValiddate)
                            dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)

                            If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                ugrow.Cells("CP DATE").Appearance.BackColor = Color.Red
                            End If

                        End If

                    Else
                        ugrow.Cells("CP DATE").Appearance.BackColor = Color.White

                    End If


                    If ugrow.Cells("TANKMODDESC").Text = "Lined Interior" Then

                        ugrow.Cells("LI INSTALL").Appearance.BackColor = Color.Yellow
                        ugrow.Cells("LI INSPECTED").Appearance.BackColor = Color.Yellow

                        dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString

                        ' install = null 10 yrs
                        If ugrow.Cells("LI INSTALL").Value Is DBNull.Value Then
                            ugrow.Cells("LI INSTALL").Appearance.BackColor = Color.Red
                            dtValiddate = DateAdd(DateInterval.Year, -10, dtValiddate)

                        Else ' if install is more than 15 yrs old, 5 yrs

                            ' first inspection = 10yrs, second and onwards = 5yrs
                            If Date.Compare(ugrow.Cells("LI INSTALL").Value, DateAdd(DateInterval.Year, -15, Today.Date)) <= 0 Then
                                dtValiddate = DateAdd(DateInterval.Year, -5, dtValiddate)
                            Else
                                dtValiddate = DateAdd(DateInterval.Year, -10, dtValiddate)
                            End If
                        End If

                        dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)

                        If ugrow.Cells("LI INSPECTED").Value Is DBNull.Value Then
                            'ugrow.Cells("LI INSPECTED").Appearance.BackColor = Color.Red
                            dttemp = CDate("01/01/0001")
                        Else
                            dttemp = ugrow.Cells("LI INSPECTED").Value
                        End If

                        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then

                            If Date.Compare(ugrow.Cells("LI INSTALL").Value, DateAdd(DateInterval.Year, -10, Today.Date)) <= 0 Then
                                ugrow.Cells("LI INSPECTED").Appearance.BackColor = Color.Red
                            End If

                        End If

                    Else

                        ugrow.Cells("LI INSTALL").Appearance.BackColor = Color.White
                        ugrow.Cells("LI INSPECTED").Appearance.BackColor = Color.White
                    End If
                End If

                'For Pipe
                Dim holdSheer As Boolean = False

                'holdSheer = fac.getFacilityCAPFieldAllowed(ugrow.Cells.Item("TANK ID").Value)


                If Not ugrow.ChildBands Is Nothing Then


                    For Each ugchildrow In ugrow.ChildBands(0).Rows   ' pipe

                        If Not ugchildrow.Cells("DateSheerValueTEST").Hidden AndAlso holdSheer Then

                            ugchildrow.Band.Columns("DateSheerValueTEST").AutoSizeMode = ColumnAutoSizeMode.VisibleRows
                            ugchildrow.Cells("DateSheerValueTEST").Hidden = True

                        End If


                        If ugchildrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then

                            If (ugchildrow.Cells("PIPE_LD").Text.ToString = "Continuous Interstitial Monitoring") Then

                                If ugchildrow.Cells("DateElectronicDeviceInspect").Value Is DBNull.Value Then
                                    ugchildrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor = Color.Red
                                    cellRed = True

                                Else

                                    dttemp = ugchildrow.Cells("DateElectronicDeviceInspect").Value
                                    dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                    dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                                    dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)

                                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                        ugchildrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor = Color.Red
                                        cellRed = True
                                    Else
                                        ugchildrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor = Color.Yellow
                                    End If

                                End If

                            Else
                                ugchildrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor = Color.White
                            End If

                        End If

                        If (ugchildrow.Cells("PIPE_TYPE_DESC").Text.ToString = "Pressurized") Then

                            If ugchildrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then

                                If Not holdSheer AndAlso ugchildrow.Cells("DateSheerValueTest").Value Is DBNull.Value Then

                                    ugchildrow.Cells("DateSheerValueTest").Appearance.BackColor = Color.Red
                                    cellRed = True

                                ElseIf Not holdSheer Then

                                    dttemp = ugchildrow.Cells("DateSheerValueTest").Value
                                    dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                    dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                                    dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)

                                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                        ugchildrow.Cells("DateSheerValueTest").Appearance.BackColor = Color.Red
                                        cellRed = True
                                    Else
                                        ugchildrow.Cells("DateSheerValueTest").Appearance.BackColor = Color.Yellow
                                    End If

                                End If

                                If (ugchildrow.Cells("PIPE_LD").Text.ToString = "Visual Interstitial Monitoring") Then

                                    If ugchildrow.Cells("DateSecondaryContainmentInspect").Value Is DBNull.Value Then
                                        ugchildrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor = Color.Red
                                        cellRed = True

                                    Else

                                        dttemp = ugchildrow.Cells("DateSecondaryContainmentInspect").Value
                                        dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                        dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                                        dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)

                                        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                            ugchildrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor = Color.Red
                                            cellRed = True
                                        Else
                                            ugchildrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor = Color.Yellow
                                        End If

                                    End If

                                Else

                                    ugchildrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor = Color.White

                                End If


                            End If

                        Else

                            If Not holdSheer Then ugchildrow.Cells("DateSheerValueTest").Appearance.BackColor = Color.White

                            If (ugchildrow.Cells("PIPE_LD").Text.ToString = "Visual Interstitial Monitoring") Then

                                If ugchildrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then

                                    If ugchildrow.Cells("DateSecondaryContainmentInspect").Value Is DBNull.Value Then

                                        ugchildrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor = Color.Red
                                        cellRed = True
                                    Else
                                        dttemp = ugchildrow.Cells("DateSecondaryContainmentInspect").Value
                                        dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                        dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                                        dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)

                                        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                            ugchildrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor = Color.Red
                                            cellRed = True
                                        Else
                                            ugchildrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor = Color.Yellow
                                        End If

                                    End If

                                Else
                                    ugchildrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor = Color.White
                                End If
                                If (ugchildrow.Cells("PIPE_LD").Text.ToString = "Continuous Interstitial Monitoring") Then
                                    If ugchildrow.Cells("DateElectronicDeviceInspect").Value Is DBNull.Value Then
                                        ugchildrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor = Color.Red
                                        cellRed = True
                                    Else
                                        dttemp = ugchildrow.Cells("DateElectronicDeviceInspect").Value
                                        dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                        dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                                        dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                                        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                            ugchildrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor = Color.Red
                                            cellRed = True
                                        Else
                                            ugchildrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor = Color.Yellow
                                        End If
                                    End If
                                End If
                            ElseIf Not (ugchildrow.Cells("PIPE_LD").Text.ToString = "Continuous Interstitial Monitoring") Then
                                ugchildrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor = Color.White
                            End If
                        End If

                        'if pipe status is ciu
                        If ugchildrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then
                            ' If ugchildrow.Cells("ALLD_TEST").Text.IndexOf("Mechanical") >= 0 Then
                            If ugchildrow.Cells("PIPE_TYPE_DESC").Text.IndexOf("Pressurized") >= 0 And ugchildrow.Cells("PIPE_LD").Text.IndexOf("Deferred") < 0 Then
                                ugchildrow.Cells("ALLD_TEST_DATE").Appearance.BackColor = Color.Yellow
                                If ugchildrow.Cells("ALLD_TEST_DATE").Value Is DBNull.Value Then
                                    ugchildrow.Cells("ALLD_TEST_DATE").Appearance.BackColor = Color.Red
                                    cellRed = True
                                Else
                                    dttemp = ugchildrow.Cells("ALLD_TEST_DATE").Value
                                    dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                    dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                                    dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                        ugchildrow.Cells("ALLD_TEST_DATE").Appearance.BackColor = Color.Red
                                        cellRed = True
                                    End If
                                End If
                            Else
                                ugchildrow.Cells("ALLD_TEST_DATE").Appearance.BackColor = Color.White
                            End If

                            If ugchildrow.Cells("PIPE_LD").Text.IndexOf("Line Tightness Testing") >= 0 Then
                                ugchildrow.Cells("TT DATE").Appearance.BackColor = Color.Yellow
                                If ugchildrow.Cells("TT DATE").Value Is DBNull.Value Then
                                    ugchildrow.Cells("TT DATE").Appearance.BackColor = Color.Red
                                Else
                                    dttemp = ugchildrow.Cells("TT DATE").Value
                                    If ugchildrow.Cells("PIPE_TYPE_DESC").Text.IndexOf("U.S. Suction") >= 0 Then
                                        dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                        dtValiddate = DateAdd(DateInterval.Year, -3, dtValiddate)
                                        dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                                    ElseIf ugchildrow.Cells("PIPE_TYPE_DESC").Text.IndexOf("Pressurized") >= 0 Then
                                        dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                        dtValiddate = DateAdd(DateInterval.Year, -1, dtValiddate)
                                        dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                                    End If
                                    If ugchildrow.Cells("PIPE_TYPE_DESC").Text.IndexOf("U.S. Suction") >= 0 Or ugchildrow.Cells("PIPE_TYPE_DESC").Text.IndexOf("Pressurized") >= 0 Then
                                        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                            ugchildrow.Cells("TT DATE").Appearance.BackColor = Color.Red
                                        End If
                                    End If
                                End If
                            Else
                                ugchildrow.Cells("TT DATE").Appearance.BackColor = Color.White
                            End If
                        End If
                        ' if pipe status is ciu / tosi
                        If ugchildrow.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Or ugchildrow.Cells("STATUS").Text.IndexOf("Temporarily Out of Service Indefinitely") >= 0 Then
                            If ugchildrow.Cells("PIPE_MOD_DESC").Text.IndexOf("Cathodically Protected") >= 0 Then
                                ugchildrow.Cells("CP DATE").Appearance.BackColor = Color.Yellow
                                If ugchildrow.Cells("CP DATE").Value Is DBNull.Value Then
                                    ugchildrow.Cells("CP DATE").Appearance.BackColor = Color.Red
                                    dttemp = CDate("01/01/0001")
                                Else
                                    dttemp = ugchildrow.Cells("CP DATE").Value
                                End If
                                dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                dtValiddate = DateAdd(DateInterval.Year, -3, dtValiddate)
                                dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                                If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                    ugchildrow.Cells("CP DATE").Appearance.BackColor = Color.Red
                                End If
                            Else
                                ugchildrow.Cells("CP DATE").Appearance.BackColor = Color.White
                            End If
                            'If ugchildrow.Cells("DISP CP TYPE").Text.IndexOf("Cathodically Protected") >= 0 Or ugchildrow.Cells("TANK CP TYPE").Text.IndexOf("Cathodically Protected") >= 0 Then
                            If ugchildrow.Cells("TERMINATION_TYPE_DISP").Text = "611" Or ugchildrow.Cells("TERMINATION_TYPE_TANK").Text = "610" Then
                                ugchildrow.Cells("TERM CP TEST").Appearance.BackColor = Color.Yellow
                                If ugchildrow.Cells("TERMINATION_CP_INSTALLED_DATE").Value Is DBNull.Value Then
                                    ugchildrow.Cells("TERMINATION_CP_INSTALLED_DATE").Appearance.BackColor = Color.Red
                                End If
                                If ugchildrow.Cells("TERM CP TEST").Value Is DBNull.Value Then
                                    ugchildrow.Cells("TERM CP TEST").Appearance.BackColor = Color.Red
                                Else
                                    dttemp = ugchildrow.Cells("TERM CP TEST").Value
                                    dtValiddate = Today.Month.ToString + "/1/" + Today.Year.ToString
                                    dtValiddate = DateAdd(DateInterval.Year, -3, dtValiddate)
                                    dtValiddate = DateAdd(DateInterval.Month, 3, dtValiddate)
                                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValiddate, dttemp) > 0 Then
                                        ugchildrow.Cells("TERM CP TEST").Appearance.BackColor = Color.Red
                                    End If
                                End If
                            Else
                                ugchildrow.Cells("TERM CP TEST").Appearance.BackColor = Color.White
                            End If
                        End If
                    Next
                End If
            Next
            ''by hua cao 01/17/09 apply CAPstatus change on new date fields
            ' If pTank.FacCapStatus = 1 Then
            'If cellRed Then
            ' Dim LocalUserSettings1 As Microsoft.Win32.Registry
            ' Dim conSQLConnection1 As New SqlConnection
            ' Dim cmdSQLCommand1 As New SqlCommand
            ' conSQLConnection1.ConnectionString = LocalUserSettings1.CurrentUser.GetValue("MusterSQLConnection")
            ' conSQLConnection1.Open()
            ' cmdSQLCommand1.Connection = conSQLConnection1
            ' cmdSQLCommand1.CommandText = "Update tblReg_Facility set CAP_Status = 0 where facility_id = " + ugrow.Cells(0).Value.ToString

            ' cmdSQLCommand1.ExecuteNonQuery()
            ' conSQLConnection1.Close()
            ' cmdSQLCommand1.Dispose()
            ' conSQLConnection1.Dispose()
            '''oOwner.Facilities.CapStatus = 0
            'pTank.FacCapStatus = 0
            'cellRed = False

            ' RaiseEvent Fillfacilities(pTank.FacilityId)


            'End If

            ' End If

            ugTankandPipe.ActiveRow = ugTankandPipe.Rows(0)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub btnCopyCurrenttoNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyCurrenttoNext.Click

        Dim baserow As UltraGridRow
        Dim basebandname As String
        Dim basebandindex As Integer
        Dim mbresult As MsgBoxResult
        Dim baseband As UltraGridChildBand
        Dim nextrow As UltraGridRow
        Dim parentindex As Integer
        Dim nextparentrow As UltraGridRow
        Dim parentrow As UltraGridRow
        Dim i As Integer = 0
        Dim StartIndex As Integer = 0
        Dim EndIndex As Integer = 0
        Dim bolLineUpdated As Boolean

        Try

            If ugTankandPipe.DataSource Is Nothing Then
                MsgBox("No Records to Copy.")
                Exit Sub
            End If

            If ugTankandPipe.ActiveRow Is Nothing Then
                MsgBox("Please Select a Row")
                Exit Sub
            End If

            bolLineUpdated = False

            baserow = ugTankandPipe.ActiveRow
            basebandname = baserow.Band.ToString
            basebandindex = baserow.Band.Index
            If Not baserow.ParentRow Is Nothing Then
                baseband = baserow.ParentRow.ChildBands(basebandname)
                StartIndex = 6
                EndIndex = 10
            Else
                baseband = Nothing
                'StartIndex = 4
                'EndIndex = 12
                StartIndex = 7
                EndIndex = 10
            End If

            If baseband Is Nothing Then
                If ugTankandPipe.Rows.Count - 1 < baserow.Index + 1 Then
                    nextrow = ugTankandPipe.Rows(0)
                Else
                    nextrow = ugTankandPipe.Rows(baserow.Index + 1)
                End If
                For i = StartIndex To EndIndex
                    'nextrow.Cells(i).Value = baserow.Cells(i).Value

                    ' Only copy fields that are updatable
                    If (baserow.Cells(i).Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        baserow.Cells(i).Appearance.BackColor.Equals(System.Drawing.Color.Red)) _
                    And (nextrow.Cells(i).Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        nextrow.Cells(i).Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                        nextrow.Cells(i).Value = baserow.Cells(i).Value
                        bolLineUpdated = True
                    End If
                Next

            Else

                parentindex = baserow.ParentRow.Index
                If baseband.Rows.Count - 1 < baserow.Index + 1 Then
                    Dim newparentindex As Integer
                    newparentindex = baserow.ParentRow.Index + 1
                    If ugTankandPipe.Rows.Count - 1 < newparentindex Then
                        nextparentrow = ugTankandPipe.Rows(0)
                    Else
                        nextparentrow = ugTankandPipe.Rows(newparentindex)
                        Do While nextparentrow.HasChild = False
                            If ugTankandPipe.Rows.Count - 1 = newparentindex + 1 Then
                                newparentindex = 0
                            End If
                            nextparentrow = ugTankandPipe.Rows(newparentindex)
                            newparentindex += 1
                        Loop
                    End If
                    baseband = nextparentrow.ChildBands(basebandname)
                    nextrow = baseband.Rows(0)
                Else
                    nextrow = baseband.Rows(baserow.Index + 1)
                End If
                For i = StartIndex To EndIndex
                    If i <> 9 Then
                        'nextrow.Cells(i).Value = baserow.Cells(i).Value

                        ' Only copy fields that are updatable
                        If (baserow.Cells(i).Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                            baserow.Cells(i).Appearance.BackColor.Equals(System.Drawing.Color.Red)) _
                        And (nextrow.Cells(i).Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                            nextrow.Cells(i).Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                            nextrow.Cells(i).Value = baserow.Cells(i).Value
                            bolLineUpdated = True
                        End If

                    End If
                Next
            End If
            If bolLineUpdated = True Then
                nextrow.CellAppearance.BackColor = Color.SkyBlue
                'Next Line Added By Elango on Feb 23 2005 
                ugTankandPipe.ActiveRow = nextrow
                ugTankandPipe.ActiveRow = baserow
                ugTankandPipe.ActiveRow = nextrow
            Else
                MessageBox.Show("No Fields Qualify For Update.", "Copy Current To Next")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnApplytoAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApplytoAll.Click

        Dim baserow As UltraGridRow
        'Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim i As Integer = 0
        Dim bolLineUpdated As Boolean

        Try
            If ugTankandPipe.DataSource Is Nothing Then
                MsgBox("No Records to Copy")
                Exit Sub
            End If

            If ugTankandPipe.ActiveRow Is Nothing Then
                MsgBox("Please Select a Row")
                Exit Sub
            End If
            baserow = ugTankandPipe.ActiveRow

            ' to prevent executing cellupdate event
            bolEvaluatingCellValue = True
            bolEvaluatingCellValueInProgress = True

            For Each ugrow In ugTankandPipe.Rows
                bolLineUpdated = False
                'For Tank
                If baserow.Band.Index = 0 Then
                    If (baserow.Cells("LI INSTALL").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        baserow.Cells("LI INSTALL").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                        (ugrow.Cells("LI INSTALL").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        ugrow.Cells("LI INSTALL").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                        ugrow.Cells("LI INSTALL").Value = baserow.Cells("LI INSTALL").Value
                        bolLineUpdated = True
                    End If
                    If (baserow.Cells("LI INSPECTED").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        baserow.Cells("LI INSPECTED").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                        (ugrow.Cells("LI INSPECTED").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        baserow.Cells("LI INSPECTED").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                        ugrow.Cells("LI INSPECTED").Value = baserow.Cells("LI INSPECTED").Value
                        bolLineUpdated = True
                    End If
                    If (baserow.Cells("TT DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        baserow.Cells("TT DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                        (ugrow.Cells("TT DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        ugrow.Cells("TT DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                        ugrow.Cells("TT DATE").Value = baserow.Cells("TT DATE").Value
                        bolLineUpdated = True
                    End If
                    If (baserow.Cells("CP DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        baserow.Cells("CP DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                        (ugrow.Cells("CP DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        ugrow.Cells("CP DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                        ugrow.Cells("CP DATE").Value = baserow.Cells("CP DATE").Value
                        bolLineUpdated = True
                    End If
                    If (baserow.Cells("DateSpillPreventionLastTested").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                       baserow.Cells("DateSpillPreventionLastTested").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                       (ugrow.Cells("DateSpillPreventionLastTested").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                       ugrow.Cells("DateSpillPreventionLastTested").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                        ugrow.Cells("DateSpillPreventionLastTested").Value = baserow.Cells("DateSpillPreventionLastTested").Value
                        bolLineUpdated = True
                    End If
                    If (baserow.Cells("DateOverfillPreventionLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                    baserow.Cells("DateOverfillPreventionLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                    (ugrow.Cells("DateOverfillPreventionLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                    ugrow.Cells("DateOverfillPreventionLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                        ugrow.Cells("DateOverfillPreventionLastInspected").Value = baserow.Cells("DateOverfillPreventionLastInspected").Value
                        bolLineUpdated = True
                    End If
                    'If (baserow.Cells("DateSecondaryContainmentLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                    '   baserow.Cells("DateSecondaryContainmentLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                    '  (ugrow.Cells("DateSecondaryContainmentLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                    '  ugrow.Cells("DateSecondaryContainmentLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                    ' ugrow.Cells("DateSecondaryContainmentLastInspected").Value = baserow.Cells("DateSecondaryContainmentLastInspected").Value
                    ' bolLineUpdated = True
                    'End If
                    If (baserow.Cells("DateElectronicDeviceInspected").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                      baserow.Cells("DateElectronicDeviceInspected").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                     (ugrow.Cells("DateElectronicDeviceInspected").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                     ugrow.Cells("DateElectronicDeviceInspected").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                        ugrow.Cells("DateElectronicDeviceInspected").Value = baserow.Cells("DateElectronicDeviceInspected").Value
                        bolLineUpdated = True
                    End If
                    If (baserow.Cells("DateATGLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        baserow.Cells("DateATGLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                       (ugrow.Cells("DateATGLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                        ugrow.Cells("DateATGLastInspected").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                        ugrow.Cells("DateATGLastInspected").Value = baserow.Cells("DateATGLastInspected").Value
                        bolLineUpdated = True
                    End If
                    If bolLineUpdated = True Then
                        ugrow.CellAppearance.BackColor = Color.SkyBlue
                        ugTankandPipe.ActiveRow = ugrow
                    End If
                End If

                'For Pipe
                If baserow.Band.Index = 1 Then
                    If Not ugrow.ChildBands Is Nothing Then
                        For Each Childrow In ugrow.ChildBands(0).Rows
                            bolLineUpdated = False
                            If (baserow.Cells("ALLD_TEST_DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                                baserow.Cells("ALLD_TEST_DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                                (Childrow.Cells("ALLD_TEST_DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                                Childrow.Cells("ALLD_TEST_DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                                Childrow.Cells("ALLD_TEST_DATE").Value = baserow.Cells("ALLD_TEST_DATE").Value
                                bolLineUpdated = True
                            End If
                            If (baserow.Cells("TT DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                                baserow.Cells("TT DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                                (Childrow.Cells("TT DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                                Childrow.Cells("TT DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                                Childrow.Cells("TT DATE").Value = baserow.Cells("TT DATE").Value
                                bolLineUpdated = True
                            End If
                            If (baserow.Cells("CP DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                                baserow.Cells("CP DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                                (Childrow.Cells("CP DATE").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                                Childrow.Cells("CP DATE").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                                Childrow.Cells("CP DATE").Value = baserow.Cells("CP DATE").Value
                                bolLineUpdated = True
                            End If
                            If (baserow.Cells("TERM CP TEST").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                                baserow.Cells("TERM CP TEST").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                                (Childrow.Cells("TERM CP TEST").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                                Childrow.Cells("TERM CP TEST").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                                Childrow.Cells("TERM CP TEST").Value = baserow.Cells("TERM CP TEST").Value
                                bolLineUpdated = True
                            End If
                            If (baserow.Cells("DateSheerValueTest").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                               baserow.Cells("DateSheerValueTest").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                               (Childrow.Cells("DateSheerValueTest").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                               Childrow.Cells("DateSheerValueTest").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                                Childrow.Cells("DateSheerValueTest").Value = baserow.Cells("DateSheerValueTest").Value
                                bolLineUpdated = True
                            End If
                            If (baserow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                              baserow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                              (Childrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                              Childrow.Cells("DateSecondaryContainmentInspect").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                                Childrow.Cells("DateSecondaryContainmentInspect").Value = baserow.Cells("DateSecondaryContainmentInspect").Value
                                bolLineUpdated = True
                            End If
                            If (baserow.Cells("DateElectronicDeviceInspect").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                            baserow.Cells("DateElectronicDeviceInspect").Appearance.BackColor.Equals(System.Drawing.Color.Red)) And _
                            (Childrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor.Equals(System.Drawing.Color.Yellow) Or _
                            Childrow.Cells("DateElectronicDeviceInspect").Appearance.BackColor.Equals(System.Drawing.Color.Red)) Then
                                Childrow.Cells("DateElectronicDeviceInspect").Value = baserow.Cells("DateElectronicDeviceInspect").Value
                                bolLineUpdated = True
                            End If

                            If bolLineUpdated = True Then
                                Childrow.CellAppearance.BackColor = Color.Pink
                                ugTankandPipe.ActiveRow = Childrow
                            End If
                        Next
                    End If
                End If
            Next
            ugTankandPipe.ActiveRow = baserow

            If Me.bolUseInspectionMode Then
                ApplyChanges()
                LoadTankandPipe(nFacilityID)
            End If

            bolEvaluatingCellValue = False

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolEvaluatingCellValue = False
            bolEvaluatingCellValueInProgress = False

        End Try
    End Sub



    Sub UpdateLabel()
        If Me.bolUseInspectionMode Then
            If isDirty Then
                Me.Label1.BackColor = Color.IndianRed
                Me.Label1.Text = "Inspection CAP data needs to be saved"
            Else
                Me.Label1.BackColor = Color.Cornsilk
                Me.Label1.Text = "Most Current CAP data for Inspection"
            End If
        Else
            If isDirty Then
                Me.Label1.BackColor = Color.IndianRed
                Me.Label1.Text = "Registration CAP data needs to be saved"
            Else
                Me.Label1.BackColor = Color.Cornsilk
                Me.Label1.Text = "Most Current CAP data in Registration"
            End If





        End If

    End Sub

    Public Function ApplyChanges(Optional ByVal row As DataRow = Nothing) As Boolean

        Try
            Dim nFacCAPStatus As Integer = 0
            Dim dtTemp As Date = CDate("01/01/0001")

            Dim drow As DataRow
            Dim TankFlag As Boolean
            Dim PipeFlag As Boolean
            Dim changes As Boolean = False


            If SaveAll Then
                row = Nothing
            End If


            'oOwner.Facilities.Retrieve(oOwner.OwnerInfo, nFacilityID, , "FACILITY", False, True)
            'pTank = oOwner.Facilities.FacilityTanks

            For Each drow In dsCAPTankandPipe.Tables(0).Rows

                If row Is Nothing OrElse drow Is row Then
                    If drow.RowState = DataRowState.Modified Or SaveAll Then


                        Dim success As Boolean = False

                        success = oOwner.Facilities.SaveTANKCAPData(Me.bolUseInspectionMode, nFacilityID, CInt(drow("TANK ID")), _
                        drow("DateSpillPreventionLastTested"), drow("DateOverfillPreventionLastInspected"), _
                         drow("DateElectronicDeviceInspected"), drow("CP DATE"), drow("LI INSPECTED"), _
                         drow("DateATGLastInspected"), drow("TT DATE"), MusterContainer.AppUser.ID, drow("LI INSTALL"), _
                         drow("DateSpillPreventionInstalled"), drow("DateOverfillPreventionInstalled"))



                        ' If Not pTank.IsDirty Then
                        ' pTank.IsDirty = True
                        'End If

                        'If pTank.IsDirty Then
                        ' bolDisplayErrmessage = True
                        ' nFacCAPStatus = oOwner.Facilities.CapStatus
                        'If pTank.TankId <= 0 Then
                        'pTank.CreatedBy = MusterContainer.AppUser.ID
                        'Else
                        '   pTank.ModifiedBy = MusterContainer.AppUser.ID
                        'End If
                        'success = pTank.Save(UIUtilsGen.ModuleID.Registration, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, , , False)
                        'If Not UIUtilsGen.HasRights(returnVal) Then
                        ' Exit Sub
                        'End If

                        If success Then
                            If Not bolFacCapStatusChanged Then
                                ' If nFacCAPStatus <> oOwner.Facilities.CapStatus Then
                                '    bolFacCapStatusChanged = True

                                '   oOwner.GetCAPParticipationLevel()
                                '  CallingForm.Tag = "1"
                                'End If
                            End If
                            changes = True
                        Else
                            MessageBox.Show("Errors in Tank Validation.", "Tank Validation")
                            Exit Function
                        End If
                        TankFlag = True
                    End If
                End If


                'End If
            Next

            'For Pipes
            For Each drow In dsCAPTankandPipe.Tables(1).Rows

                If row Is Nothing OrElse drow Is row Then

                    If drow.RowState = DataRowState.Modified Or SaveAll Then


                        Dim success As Boolean = False

                        success = oOwner.Facilities.SavePIPECAPData(Me.bolUseInspectionMode, nFacilityID, CInt(IIf(drow("PIPE ID") Is DBNull.Value, 0, drow("PIPE ID"))), CInt(IIf(drow("TANK ID") Is DBNull.Value, 0, drow("TANK ID"))), _
                          drow("ALLD_TEST_DATE"), drow("TT DATE"), drow("CP DATE"), drow("TERM CP TEST"), drow("DateSheerValueTest"), _
                           drow("DateSecondaryContainmentInspect"), drow("DateElectronicDeviceInspect"), MusterContainer.AppUser.ID)



                        ' If Not pPipe.IsDirty Then
                        'pPipe.IsDirty = True
                        'End If
                        'If pPipe.IsDirty Then
                        ' bolDisplayErrmessage = True

                        ' nFacCAPStatus = oOwner.Facilities.CapStatus

                        'If pPipe.PipeID <= 0 Then
                        'pPipe.CreatedBy = MusterContainer.AppUser.ID
                        'Else
                        'pPipe.ModifiedBy = MusterContainer.AppUser.ID
                        'End If
                        'success = pPipe.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, , False)
                        ' If Not UIUtilsGen.HasRights(returnVal) Then
                        '    Exit Sub
                        'End If

                        If success Then
                            If Not bolFacCapStatusChanged Then
                                '  If nFacCAPStatus <> pTank.FacCapStatus Then
                                '  bolFacCapStatusChanged = True
                                ' pTank.FacCapStatus = pPipe.FacCapStatus
                                'oOwner.Facilities.CapStatusOriginal = pPipe.FacCapStatus
                                'oOwner.Facilities.CapStatus = pPipe.FacCapStatus

                                'oOwner.GetCAPParticipationLevel()
                                'CallingForm.Tag = "1"
                                'End If
                            End If
                            changes = True

                        Else
                            MessageBox.Show("Errors in Pipe Validation.", "Pipe Validation")
                            Exit Function
                        End If

                        PipeFlag = True
                    End If

                End If
                'End If
            Next

            SaveAll = False

            If TankFlag = True And PipeFlag = True Then
                MessageBox.Show("Tank(s) and Pipe(s) Modified Successfully.", "Modify Tank(s) and Pipe(s)")
            ElseIf TankFlag = True And PipeFlag = False Then
                MessageBox.Show("Tank(s) Modified Successfully.", "Modify Tank(s)")
            ElseIf TankFlag = False And PipeFlag = True Then
                MessageBox.Show("Pipe(s) Modified Successfully.", "Modify Pipe(s)")
            ElseIf TankFlag = False And PipeFlag = False Then
                MessageBox.Show("No Modified Records to Update.", "Modify Tank(s) and Pipe(s)")
            End If

            isDirty = False
            UpdateLabel()



            RaiseEvent CapChanged()

            Return changes

            ' Refresh Tank and Pipe Grid here
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Sub ugTankandPipe_AfterEnterEditMode(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugTankandPipe.AfterEnterEditMode
        Try
            If Not (ugTankandPipe.ActiveCell.Appearance.BackColor.Equals(Color.Yellow) Or ugTankandPipe.ActiveCell.Appearance.BackColor.Equals(Color.Red)) Then
                ugTankandPipe.PerformAction(UltraGridAction.ExitEditMode, False, False)
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub ugTankandPipe_AfterCellListCloseUp(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTankandPipe.AfterCellListCloseUp
    '    Dim cellType As System.Type
    '    Dim strvalue As String
    '    Dim dttemp As Date
    '    cellType = e.Cell.Value.GetType

    '    Exit Sub

    '    Try
    '        'If cellType.Equals(GetType(Date)) Or cellType.Equals(GetType(DBNull)) And IsDate(e.Cell.Text) Then
    '        'Get the Selected Cell Value
    '        'End If
    '        ugTankandPipe_AfterCellUpdate(sender, e)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    'Private Sub ugTankandPipe_AfterCellUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTankandPipe.AfterCellUpdate

    '    Dim cellType As System.Type
    '    Dim strvalue As String

    '    If bolEvaluatingCellValue Then
    '        Exit Sub
    '    End If
    '    'bolEvaluatingCellValue = True

    '    'ValidateCells(e)

    '    'bolEvaluatingCellValue = False

    'End Sub


#End Region

#Region "External Events"


    Private Sub pTank_evtTankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String) Handles pTank.evtTankValidationErr
        If Not Me.bolDisplayErrmessage Then
            Me.bolDisplayErrmessage = True
            Exit Sub
        End If

        If strMessage <> String.Empty And MsgBox(strMessage) = MsgBoxResult.OK Then
            strMessage = String.Empty
        End If
        Me.bolDisplayErrmessage = False
        bolValidateSuccess = False
    End Sub

    Private Sub pPipe_evtPipeErr(ByVal StrMessage As String) Handles pPipe.evtPipeErr
        If Not Me.bolDisplayErrmessage Then
            Me.bolDisplayErrmessage = True
            Exit Sub
        End If
        If StrMessage <> String.Empty And MsgBox(StrMessage) = MsgBoxResult.OK Then
            StrMessage = String.Empty
        End If
        bolValidateSuccess = False
        Me.bolDisplayErrmessage = False
    End Sub

#End Region

    Private Function GetUGRow(ByVal tankID As Integer, Optional ByVal compNum As Integer = 0, Optional ByVal pipeID As Integer = 0) As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing
        Try
            If tankID > 0 And compNum = 0 And pipeID = 0 Then
                For Each drTank As Infragistics.Win.UltraWinGrid.UltraGridRow In ugTankandPipe.Rows
                    If drTank.Cells("TANK ID").Value = tankID Then
                        ugRow = drTank
                        Exit Try
                    End If
                Next
            ElseIf tankID > 0 And compNum > 0 And pipeID > 0 Then
                For Each drTank As Infragistics.Win.UltraWinGrid.UltraGridRow In ugTankandPipe.Rows
                    If drTank.Cells("TANK ID").Value = tankID Then
                        If Not drTank.ChildBands Is Nothing Then
                            If Not drTank.ChildBands(0).Rows Is Nothing Then
                                For Each drPipe As Infragistics.Win.UltraWinGrid.UltraGridRow In drTank.ChildBands(0).Rows
                                    If drPipe.Cells("PIPE ID").Value = pipeID And drPipe.Cells("COMPARTMENT_NUMBER").Value = compNum Then
                                        ugRow = drPipe
                                        Exit Try
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return ugRow
    End Function


    Public Function ValidateUG() As Boolean
        Dim strMsg As String = String.Empty
        Dim str As String = String.Empty
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow

        ' 2197 validate cap fields only if facility is cap candidate
        If CAP_STATUS Then
            For Each drow As DataRow In dsCAPTankandPipe.Tables(0).Rows
                If drow.RowState = DataRowState.Modified Then
                    ugRow = GetUGRow(drow("TANK ID"))
                    If Not ugRow Is Nothing Then
                        For Each ugCell As Infragistics.Win.UltraWinGrid.UltraGridCell In ugRow.Cells
                            If ugcell.Column.Key.ToUpper = "CP DATE" Or _
                                ugcell.Column.Key.ToUpper = "TERM CP TEST" Or _
                                ugcell.Column.Key.ToUpper = "TT DATE" Or _
                                ugcell.Column.Key.ToUpper = "LI INSPECTED" Or _
                                ugcell.Column.Key.ToUpper = "LI INSTALL" Then
                                str = String.Empty
                                str = ValidateCells(ugcell)
                                If str <> String.Empty Then
                                    strMsg += str + vbCrLf
                                End If
                            End If
                        Next
                    End If
                End If
            Next

            For Each drow As DataRow In dsCAPTankandPipe.Tables(1).Rows
                If drow.RowState = DataRowState.Modified Then
                    ugRow = GetUGRow(drow("TANK ID"), drow("COMPARTMENT_NUMBER"), drow("PIPE ID"))
                    If Not ugRow Is Nothing Then
                        For Each ugCell As Infragistics.Win.UltraWinGrid.UltraGridCell In ugRow.Cells
                            If ugcell.Column.Key.ToUpper = "CP DATE" Or _
                                ugcell.Column.Key.ToUpper = "ALLD_TEST_DATE" Or _
                                ugcell.Column.Key.ToUpper = "TT DATE" Or _
                                ugcell.Column.Key.ToUpper = "TERM CP TEST" Then
                                str = String.Empty
                                str = ValidateCells(ugcell)
                                If str <> String.Empty Then
                                    strMsg += str + vbCrLf
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End If

        If strMsg <> String.Empty Then
            Dim mb As MessageBoxCustom
            If mb.Show("Invalid dates entered for CAP." + vbCrLf + _
                    "Valid dates for" + vbCrLf + _
                    strMsg + vbCrLf + _
                    "Do you want to continue?", "CAP Validation", MessageBoxButtons.YesNo) = DialogResult.No Then
                Return False
            Else
                Return True
            End If
        Else
            Return True
        End If
    End Function

    'Private Sub ValidateCells(ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)
    Private Function ValidateCells(ByVal cell As Infragistics.Win.UltraWinGrid.UltraGridCell) As String
        Dim cellType As System.Type
        Dim strvalue As String = String.Empty
        Dim dttemp, dtValidDate As Date
        Dim dtTodayPlus90Days As Date = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 4, CDate(Today.Month.ToString + "/1/" + Today.Year.ToString)))
        Dim dt As Date

        cellType = cell.Value.GetType

        If cellType.Equals(GetType(Date)) Or cellType.Equals(GetType(DBNull)) And IsDate(cell.Text) Then
            Dim aDate As Date = CDate(cell.Text)
            If cell.Row.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Or cell.Row.Cells("STATUS").Text.IndexOf("Temporarily Out of Service Indefinitely") >= 0 Then
                Dim isValiddtTempDate As Boolean = False
                If cell.Band.Index = 0 Then
                    ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                    ' -+-+-+ Tank Validations  +-+-+-
                    ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

                    Today.Month.ToString(+"/1/" + Today.Year.ToString)

                    Select Case cell.Column.Key.ToUpper
                        Case "CP DATE", "TERM CP TEST"
                            dttemp = aDate
                            dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                            dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                            If cell.Row.Cells("TCPINSTALLDATE").Value Is DBNull.Value Then
                                dt = CDate("01/01/0001")
                            Else
                                dt = cell.Row.Cells("TCPINSTALLDATE").Value
                            End If
                            If Date.Compare(dt, dtValidDate) > 0 Then
                                dtValidDate = dt
                            End If
                            isValiddtTempDate = True
                        Case "TT DATE"
                            If cell.Row.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then
                                dttemp = aDate
                                dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                                dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                isValiddtTempDate = True
                            End If
                        Case "LI INSPECTED"
                            dttemp = aDate
                            dtValidDate = Today

                            ' install = null 10 yrs
                            If cell.Row.Cells("LI INSTALL").Value Is DBNull.Value Then
                                dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                            ElseIf Date.Compare(cell.Row.Cells("LI INSTALL").Value, CDate("01/01/0001")) = 0 Then
                                dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                            Else ' if install is more than 15 yrs old, 5 yrs
                                ' first inspection = 10yrs, second and onwards = 5yrs
                                If Date.Compare(cell.Row.Cells("LI INSTALL").Value, DateAdd(DateInterval.Year, -15, Today.Date)) <= 0 Then
                                    dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                                Else
                                    dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                                End If
                            End If

                            'If cell.OriginalValue Is DBNull.Value Then
                            '    dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                            'Else
                            '    dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                            'End If

                            ' dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                            If cell.Row.Cells("LI INSTALL").Value Is DBNull.Value Then
                                dt = CDate("01/01/0001")
                            Else
                                dt = cell.Row.Cells("LI INSTALL").Value
                            End If
                            If Date.Compare(dt, dtValidDate) > 0 Then
                                dtValidDate = dt
                            End If
                            isValiddtTempDate = True
                        Case "LI INSTALL"
                            dttemp = aDate
                            dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                            dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                            isValiddtTempDate = True
                    End Select
                    If isValiddtTempDate Then
                        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                            strvalue = cell.Column.ToString + " : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString + vbCrLf + " for Tank Site ID : " + cell.Row.Cells("TANK SITE ID").Text
                            'If MsgBox(cell.Column.ToString + " must be greater than or equal to " + dtValidDate.ToShortDateString + vbCrLf + " and less than or equal to  - " + dtTodayPlus90Days.ToShortDateString + vbCrLf + " for Tank Site ID : " + cell.Row.Cells("TANK SITE ID").Text + " in Facility : " + cell.Row.Cells("FACILITY_ID").Text + vbCrLf + "Do you want to continue?", MsgBoxStyle.YesNo, "CAP Validation") = MsgBoxResult.No Then
                            '    If Not IsDBNull(cell.OriginalValue) Then
                            '        cell.Value = cell.OriginalValue
                            '    Else
                            '        cell.Value = System.DBNull.Value
                            '    End If
                            'End If
                        End If
                    End If
                ElseIf cell.Band.Index = 1 Then
                    ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                    ' -+-+-+ Pipe Validations  +-+-+-
                    ' -+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                    dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString

                    Select Case cell.Column.Key.ToUpper
                        Case "CP DATE"
                            dttemp = aDate
                            dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                            dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                            If cell.Row.Cells("PIPE_CP_INSTALLED_DATE").Value Is DBNull.Value Then
                                dt = CDate("01/01/0001")
                            Else
                                dt = cell.Row.Cells("PIPE_CP_INSTALLED_DATE").Value
                            End If
                            If Date.Compare(dt, dtValidDate) > 0 Then
                                dtValidDate = dt
                            End If
                            isValiddtTempDate = True
                        Case "ALLD_TEST_DATE"
                            If cell.Row.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then
                                dttemp = aDate
                                dtValidDate = DateAdd(DateInterval.Year, -1, dtValidDate)
                                dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                isValiddtTempDate = True
                            End If
                        Case "TT DATE"
                            If cell.Row.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Then
                                If cell.Row.Cells("PIPE_TYPE_DESC").Text.IndexOf("U.S.Suction") >= 0 Then
                                    dttemp = aDate
                                    dtValidDate = Today
                                    dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                    isValiddtTempDate = True
                                ElseIf cell.Row.Cells("PIPE_TYPE_DESC").Text.IndexOf("Pressurized") >= 0 Then
                                    dttemp = aDate
                                    dtValidDate = DateAdd(DateInterval.Year, -1, dtValidDate)
                                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                    isValiddtTempDate = True
                                End If
                            End If
                        Case "TERM CP TEST"
                            If cell.Row.Cells("STATUS").Text.IndexOf("Currently In Use") >= 0 Or cell.Row.Cells("STATUS").Text.IndexOf("Temporarily Out of Service Indefinitely") >= 0 Then
                                If cell.Row.Cells("TERMINATION_TYPE_DISP").Text = "611" Or cell.Row.Cells("TERMINATION_TYPE_TANK").Text = "610" Then
                                    dttemp = aDate
                                    dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                                    If cell.Row.Cells("TERMINATION_CP_INSTALLED_DATE").Value Is DBNull.Value Then
                                        dt = CDate("01/01/0001")
                                    Else
                                        dt = cell.Row.Cells("TERMINATION_CP_INSTALLED_DATE").Value
                                    End If
                                    If Date.Compare(dt, dtValidDate) > 0 Then
                                        dtValidDate = dt
                                    End If
                                    isValiddtTempDate = True
                                End If
                            End If
                    End Select
                    If isValiddtTempDate Then
                        If Date.Compare(dttemp, Today) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                            strvalue = cell.Column.ToString + " : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString + vbCrLf + " for Pipe Site ID : " + cell.Row.Cells("PIPE SITE ID").Text + " in Tank Site ID : " + cell.Row.ParentRow.Cells("TANK SITE ID").Text
                            'If MsgBox(cell.Column.ToString + " must be greater than or equal to " + dtValidDate.ToShortDateString + vbCrLf + " and less than or equal to  - " + dtTodayPlus90Days.ToShortDateString + vbCrLf + " for Pipe Site ID : " + cell.Row.Cells("PIPE SITE ID").Text + " belonging to Tank : " + cell.Row.ParentRow.Cells("TANK SITE ID").Text + " in Facility : " + cell.Row.Cells("FACILITY_ID").Text + vbCrLf + "Do you want to continue?", MsgBoxStyle.YesNo, "CAP Validation") = MsgBoxResult.No Then
                            '    If Not IsDBNull(cell.OriginalValue) Then
                            '        cell.Value = cell.OriginalValue
                            '    Else
                            '        cell.Value = System.DBNull.Value
                            '    End If
                            'End If
                        End If
                    End If
                End If
            End If

        End If
        Return strvalue
    End Function


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub btnApplyHover(ByVal sender As Object, ByVal e As MouseEventArgs) Handles btnApplytoAll.MouseMove
        bolEvaluatingCellValue = True
    End Sub

    Private Sub btnApplyEnter(ByVal sender As Object, ByVal e As EventArgs) Handles btnApplytoAll.MouseEnter
        bolEvaluatingCellValue = True
    End Sub


    Private Sub btnApplyLeave(ByVal sender As Object, ByVal e As EventArgs) Handles btnApplytoAll.MouseLeave
        If Not Me.bolEvaluatingCellValueInProgress Then
            bolEvaluatingCellValue = False
        End If

    End Sub

    Private Sub ugTankandPipe_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugTankandPipe.AfterRowUpdate

        isDirty = True
        UpdateLabel()

        If Not bolEvaluatingCellValue Then
            Dim dr As DataRow

            If e.Row.Band.Index = 0 Then
                dr = Me.dsCAPTankandPipe.Tables(0).Select(String.Format("[TANK ID] = {0}", e.Row.Cells("TANK ID").Value))(0)

            Else
                dr = Me.dsCAPTankandPipe.Tables(1).Select(String.Format("[PIPE ID] = {0}", e.Row.Cells("PIPE ID").Value))(0)
            End If


            If Not Me.bolUseInspectionMode Then
                ApplyBusinessRules()
            End If

        End If



    End Sub



    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim bolSetActive As Boolean = False

        Try
            If ugTankandPipe.DataSource Is Nothing Then
                MsgBox("No Records to Save")
                Exit Sub
            Else
                If Not ValidateUG() Then
                    Exit Sub
                End If
            End If


            ApplyChanges()

            RaiseEvent Fillfacilities(nFacilityID)

            CallingForm.Tag = "1"

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        ''Added By Elango on Feb 10 2005.
        'Tank information..
        Dim msgResult As String


        Try
            If oOwner.Facilities.FacilityTanks.colIsDirty Then
                msgResult = MsgBox("You have changed the Tank data in this page/main Tank Page.Do you want to save the data ?", MsgBoxStyle.YesNo, "Tank")
                If msgResult = MsgBoxResult.No Then
                    oOwner.Facilities.FacilityTanks.ResetCollection()
                Else
                    oOwner.Facilities.FacilityTanks.Flush(CType(UIUtilsGen.ModuleID.Registration, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    RaiseEvent Fillfacilities(nFacilityID)
                End If
            End If

            'Pipe information
            If oOwner.Facilities.FacilityTanks.Compartments.Pipes.colIsDirty Then
                msgResult = MsgBox("You have changed the Pipe data in this page/main Pipe Page.Do you want to save the data ?", MsgBoxStyle.YesNo, "Tank")
                If msgResult = MsgBoxResult.No Then
                    oOwner.Facilities.FacilityTanks.Compartments.Pipes.ResetCollection()
                Else
                    oOwner.Facilities.FacilityTanks.Compartments.Pipes.Flush(CType(UIUtilsGen.ModuleID.Registration, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    RaiseEvent Fillfacilities(nFacilityID)

                End If
            End If

            ' Reload Tank and Pipe Here
            LoadTankandPipe(nFacilityID)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnApplytoAll_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApplytoAll.Disposed

        If Not Me.dctCapFields Is Nothing Then
            Me.dctCapFields.Clear()
        End If

        Me.dctCapFields = Nothing

    End Sub
End Class
