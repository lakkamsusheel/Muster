Public Class FacilityComplianceEvent
    Inherits System.Windows.Forms.Form
#Region "Private User Variables"
    Friend CallingForm As Form
    Private nOwnerID As Int32 = 0
    Private strOwnerName As String = String.Empty
    Private nFacID As Int64 = 0
    Private nFCEID As Integer = 0
    Private bolLoading As Boolean = False
    Private nInspectionID As Integer = 0
    Private alCitationPenalty As ArrayList
    Private alDiscrepText As ArrayList
    Private dsCitations As DataSet
    Private dtCitations As DataTable
    Private dtDiscrepText As DataTable
    Private bolcitationModified As Boolean

    Private WithEvents pFCE As New MUSTER.BusinessLogic.pFacilityComplianceEvent
    Private WithEvents frmCitation As CitationList
    Private oInspectionCitation As New MUSTER.BusinessLogic.pInspectionCitation
    Private oInspectionDiscrep As New MUSTER.BusinessLogic.pInspectionDiscrep
    Private oInspection As New MUSTER.BusinessLogic.pInspection
    Dim returnVal As String = String.Empty

    'Private WithEvents pCitPenalty As MUSTER.BusinessLogic.pCitationPenalty
    'Private pLCE As New MUSTER.BusinessLogic.pLicenseeComplianceEvent
    'Private citationInfo As MUSTER.Info.InspectionCitationInfo
    'Private objFCEInfo As MUSTER.Info.FacilityComplianceEventInfo
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByVal OwnerID As Integer = 0, Optional ByVal inspID As Integer = 0, Optional ByVal facID As Int64 = 0, Optional ByVal FCEID As Int64 = 0, Optional ByVal ownName As String = "")
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        nOwnerID = OwnerID
        nInspectionID = inspID
        nFacID = facID
        nFCEID = FCEID
        strOwnerName = ownName
        alCitationPenalty = New ArrayList
        alDiscrepText = New ArrayList
        dsCitations = New DataSet
        dtCitations = New DataTable
        dtDiscrepText = New DataTable
        bolcitationModified = False
        ' if fceid = 0, it creates a new instance
        If nFCEID <= 0 Then
            pFCE.Retrieve(0)
            pFCE.InspectionID = nInspectionID
            pFCE.OwnerID = nOwnerID
            pFCE.FacilityID = nFacID
            pFCE.Source = "ADMIN"
            pFCE.FCEDate = Today.Date
        Else
            pFCE.Retrieve(nFCEID, nInspectionID, nOwnerID, nFacID)
        End If
    End Sub
    'Public Sub New(ByVal OwnerID As Int32, ByVal inspID As Int64, Optional ByVal OwnerName As String = "", Optional ByVal facID As Int64 = 0, Optional ByVal FCEID As Int64 = 0)
    '    MyBase.New()
    '    bolLoading = True
    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()

    '    'Add any initialization after the InitializeComponent() call
    '    nOwnerID = OwnerID
    '    nInspectionID = inspID
    '    strOwnerName = OwnerName
    '    nFacID = facID
    '    nFCEID = FCEID
    '    objFCE = New MUSTER.BusinessLogic.pFacilityComplianceEvent
    '    If nFCEID <> 0 Then
    '        objFCEInfo = objFCE.Retrieve(FCEID)
    '    Else
    '        objFCEInfo = New MUSTER.Info.FacilityComplianceEventInfo(nFCEID, _
    '                                            nInspectionID, _
    '                                            nOwnerID, _
    '                                            nFacID, _
    '                                            CDate("01/01/0001"), _
    '                                            2, _
    '                                            CDate("01/01/0001"), _
    '                                            CDate("01/01/0001"), _
    '                                            MusterContainer.AppUser.ID, _
    '                                            Now, _
    '                                            MusterContainer.AppUser.ID, _
    '                                            CDate("01/01/0001"), _
    '                                            False, _
    '                                            False)
    '    End If
    '    pCitPenalty = New MUSTER.BusinessLogic.pCitationPenalty
    '    bolLoading = False
    'End Sub
    'Form overrides dispose to clean up the component list.
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
    Friend WithEvents btnCitationDelete As System.Windows.Forms.Button
    Friend WithEvents btnCitationAdd As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents chkSearchforOwner As System.Windows.Forms.CheckBox
    Friend WithEvents cmbFacility As System.Windows.Forms.ComboBox
    Friend WithEvents lblFacility As System.Windows.Forms.Label
    Public WithEvents dtComplianceEventDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblExceptionGrantedDate As System.Windows.Forms.Label
    Friend WithEvents chkSelectedOwner As System.Windows.Forms.CheckBox
    Friend WithEvents txtSelectedOwner As System.Windows.Forms.TextBox
    Friend WithEvents ugCitations As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmbSearchForOwner As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.chkSelectedOwner = New System.Windows.Forms.CheckBox
        Me.ugCitations = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.txtSelectedOwner = New System.Windows.Forms.TextBox
        Me.chkSearchforOwner = New System.Windows.Forms.CheckBox
        Me.btnCitationDelete = New System.Windows.Forms.Button
        Me.btnCitationAdd = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.cmbFacility = New System.Windows.Forms.ComboBox
        Me.lblFacility = New System.Windows.Forms.Label
        Me.dtComplianceEventDate = New System.Windows.Forms.DateTimePicker
        Me.lblExceptionGrantedDate = New System.Windows.Forms.Label
        Me.cmbSearchForOwner = New System.Windows.Forms.ComboBox
        CType(Me.ugCitations, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkSelectedOwner
        '
        Me.chkSelectedOwner.Checked = True
        Me.chkSelectedOwner.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSelectedOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSelectedOwner.Location = New System.Drawing.Point(24, 24)
        Me.chkSelectedOwner.Name = "chkSelectedOwner"
        Me.chkSelectedOwner.Size = New System.Drawing.Size(112, 16)
        Me.chkSelectedOwner.TabIndex = 0
        Me.chkSelectedOwner.Tag = "0"
        Me.chkSelectedOwner.Text = "Selected Owner"
        '
        'ugCitations
        '
        Me.ugCitations.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCitations.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugCitations.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCitations.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugCitations.Location = New System.Drawing.Point(16, 160)
        Me.ugCitations.Name = "ugCitations"
        Me.ugCitations.Size = New System.Drawing.Size(448, 152)
        Me.ugCitations.TabIndex = 5
        Me.ugCitations.Text = "Citations"
        '
        'txtSelectedOwner
        '
        Me.txtSelectedOwner.Location = New System.Drawing.Point(152, 24)
        Me.txtSelectedOwner.Name = "txtSelectedOwner"
        Me.txtSelectedOwner.ReadOnly = True
        Me.txtSelectedOwner.Size = New System.Drawing.Size(232, 20)
        Me.txtSelectedOwner.TabIndex = 1001
        Me.txtSelectedOwner.Text = ""
        '
        'chkSearchforOwner
        '
        Me.chkSearchforOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSearchforOwner.Location = New System.Drawing.Point(24, 56)
        Me.chkSearchforOwner.Name = "chkSearchforOwner"
        Me.chkSearchforOwner.Size = New System.Drawing.Size(120, 16)
        Me.chkSearchforOwner.TabIndex = 1
        Me.chkSearchforOwner.Tag = "644"
        Me.chkSearchforOwner.Text = "Search for Owner"
        '
        'btnCitationDelete
        '
        Me.btnCitationDelete.Location = New System.Drawing.Point(112, 320)
        Me.btnCitationDelete.Name = "btnCitationDelete"
        Me.btnCitationDelete.Size = New System.Drawing.Size(144, 23)
        Me.btnCitationDelete.TabIndex = 7
        Me.btnCitationDelete.Text = "Delete Citation / Discrep"
        '
        'btnCitationAdd
        '
        Me.btnCitationAdd.Location = New System.Drawing.Point(16, 320)
        Me.btnCitationAdd.Name = "btnCitationAdd"
        Me.btnCitationAdd.Size = New System.Drawing.Size(88, 23)
        Me.btnCitationAdd.TabIndex = 6
        Me.btnCitationAdd.Text = "Add Citation"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(184, 368)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 23)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(88, 368)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(88, 23)
        Me.btnSave.TabIndex = 8
        Me.btnSave.Text = "Save"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(280, 368)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(88, 23)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        '
        'cmbFacility
        '
        Me.cmbFacility.Location = New System.Drawing.Point(152, 88)
        Me.cmbFacility.Name = "cmbFacility"
        Me.cmbFacility.Size = New System.Drawing.Size(232, 21)
        Me.cmbFacility.TabIndex = 3
        '
        'lblFacility
        '
        Me.lblFacility.Location = New System.Drawing.Point(88, 88)
        Me.lblFacility.Name = "lblFacility"
        Me.lblFacility.Size = New System.Drawing.Size(48, 16)
        Me.lblFacility.TabIndex = 24
        Me.lblFacility.Text = "Facility:"
        Me.lblFacility.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtComplianceEventDate
        '
        Me.dtComplianceEventDate.Checked = False
        Me.dtComplianceEventDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtComplianceEventDate.Location = New System.Drawing.Point(152, 120)
        Me.dtComplianceEventDate.Name = "dtComplianceEventDate"
        Me.dtComplianceEventDate.ShowCheckBox = True
        Me.dtComplianceEventDate.Size = New System.Drawing.Size(104, 20)
        Me.dtComplianceEventDate.TabIndex = 4
        '
        'lblExceptionGrantedDate
        '
        Me.lblExceptionGrantedDate.Location = New System.Drawing.Point(16, 120)
        Me.lblExceptionGrantedDate.Name = "lblExceptionGrantedDate"
        Me.lblExceptionGrantedDate.Size = New System.Drawing.Size(128, 16)
        Me.lblExceptionGrantedDate.TabIndex = 247
        Me.lblExceptionGrantedDate.Text = "Compliance Event Date:"
        Me.lblExceptionGrantedDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbSearchForOwner
        '
        Me.cmbSearchForOwner.Location = New System.Drawing.Point(152, 56)
        Me.cmbSearchForOwner.Name = "cmbSearchForOwner"
        Me.cmbSearchForOwner.Size = New System.Drawing.Size(232, 21)
        Me.cmbSearchForOwner.TabIndex = 2
        '
        'FacilityComplianceEvent
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(480, 414)
        Me.Controls.Add(Me.cmbSearchForOwner)
        Me.Controls.Add(Me.dtComplianceEventDate)
        Me.Controls.Add(Me.txtSelectedOwner)
        Me.Controls.Add(Me.chkSelectedOwner)
        Me.Controls.Add(Me.lblExceptionGrantedDate)
        Me.Controls.Add(Me.cmbFacility)
        Me.Controls.Add(Me.lblFacility)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnCitationDelete)
        Me.Controls.Add(Me.btnCitationAdd)
        Me.Controls.Add(Me.chkSearchforOwner)
        Me.Controls.Add(Me.ugCitations)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FacilityComplianceEvent"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Facility Compliance Event"
        CType(Me.ugCitations, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "UI Support Routines"
    Private Sub Populate()
        Try
            ' if fceid is 0, mode = add, else edit
            ' if owner name is provided, check selected owner else check search for owner
            If strOwnerName <> String.Empty Then
                txtSelectedOwner.Text = strOwnerName
                chkSelectedOwner.Checked = True
                chkSearchforOwner.Checked = False
                cmbSearchForOwner.DataSource = Nothing
                cmbSearchForOwner.Enabled = False
            Else
                txtSelectedOwner.Text = String.Empty
                chkSelectedOwner.Checked = False
                chkSearchforOwner.Checked = True
                cmbSearchForOwner.Enabled = True
                PopulateOwners(nOwnerID)
            End If

            UIUtilsGen.SetDatePickerValue(dtComplianceEventDate, pFCE.FCEDate)

            PopulateFacility(nOwnerID, nFacID)

            PopulateCitations()

            EnableSave(pFCE.IsDirty)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateOwners(Optional ByVal ownID As Integer = 0)
        Try
            cmbSearchForOwner.DataSource = pFCE.GetOwners.Tables(0)
            cmbSearchForOwner.DisplayMember = "o_name"
            cmbSearchForOwner.ValueMember = "o_id"

            pFCE.OwnerID = ownID

            If ownID > 0 Then
                UIUtilsGen.SetComboboxItemByValue(cmbSearchForOwner, ownID)
            Else
                cmbSearchForOwner.SelectedIndex = -1
                If cmbSearchForOwner.SelectedIndex <> -1 Then
                    cmbSearchForOwner.SelectedIndex = -1
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateFacility(ByVal OwnerID As Integer, Optional ByVal facID As Integer = 0)
        Try
            cmbFacility.DataSource = pFCE.GetFacilities(OwnerID).Tables(0)
            cmbFacility.DisplayMember = "FACILITY"
            cmbFacility.ValueMember = "FACILITY_ID"

            pFCE.FacilityID = facID

            If facID > 0 Then
                UIUtilsGen.SetComboboxItemByValue(cmbFacility, facID)
            Else
                cmbFacility.SelectedIndex = -1
                If cmbFacility.SelectedIndex <> -1 Then
                    cmbFacility.SelectedIndex = -1
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateCitations()
        Dim strWhere As String = ""
        Dim str As String = String.Empty
        Try
            For Each citID As Integer In alCitationPenalty
                strWhere += " CITATION_ID = " + citID.ToString + " OR"
            Next
            If strWhere.Length > 0 Then
                strWhere = " WHERE" + strWhere.Substring(0, strWhere.Length - 2)
                dsCitations = pFCE.GetCitations(strWhere)
                PopulateDiscrepText()
                ugCitations.DataSource = dsCitations
            Else
                If nFCEID <= 0 Then
                    ugCitations.DataSource = Nothing
                Else
                    strWhere = " WHERE CITATION_ID IN (SELECT DISTINCT CITATION_ID FROM tblINS_INSPECTION_CITATION WHERE DELETED = 0 AND OCE_ID IS NULL AND FCE_ID = " + nFCEID.ToString + ")" + _
                                " AND DELETED = 0 ORDER BY CATEGORY"
                    alCitationPenalty = New ArrayList
                    alDiscrepText = New ArrayList
                    dsCitations = pFCE.GetCitations(strWhere)
                    PopulateDiscrepText()
                    ugCitations.DataSource = dsCitations
                    For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCitations.Rows
                        alCitationPenalty.Add(CType(ugRow.Cells("CITATION_ID").Value, Integer))
                        If Not ugRow.ChildBands Is Nothing Then
                            For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugRow.ChildBands(0).Rows
                                alDiscrepText.Add(ugChildRow.Cells("DISCREP TEXT").Text)
                            Next
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateDiscrepText()
        Dim strWhere As String = ""
        Dim str As String = String.Empty
        Try
            If Not dsCitations.Tables(0).Select("CATEGORY = 'DISCREPANCY'") Is Nothing Then
                If dsCitations.Tables(0).Select("CATEGORY = 'DISCREPANCY'").Length > 0 Then
                    For Each strText As String In alDiscrepText
                        strWhere += " DISCREP_TEXT = '" + strText + "' OR"
                    Next
                    If strWhere.Length > 0 Then
                        strWhere = " AND" + strWhere.Substring(0, strWhere.Length - 2)
                        Dim dt As DataTable = pFCE.GetDiscrepText(strWhere).Tables(0)
                        If Not dt Is Nothing Then
                            dsCitations.Tables.Add("DISCREPTEXT")
                            For Each dCol As DataColumn In dt.Columns
                                dsCitations.Tables("DISCREPTEXT").Columns.Add(dCol.ColumnName, dCol.DataType)
                            Next
                            Dim dr As DataRow
                            For Each dRow As DataRow In dt.Rows
                                dr = dsCitations.Tables("DISCREPTEXT").NewRow
                                For Each dCol As DataColumn In dt.Columns
                                    dr(dcol.ColumnName) = dRow.Item(dcol.ColumnName)
                                Next
                                dsCitations.Tables("DISCREPTEXT").Rows.Add(dr)
                            Next
                            AddDiscrepTextRelation()
                        End If
                    Else
                        If nFCEID > 0 Then
                            strWhere = " AND DISCREP_TEXT IN (SELECT DISTINCT [DESCRIPTION] FROM tblINS_INSPECTION_DISCREP WHERE DELETED = 0 AND INSPECTION_ID = " + nInspectionID.ToString + ")" + _
                                        " ORDER BY DISCREP_TEXT"
                            alDiscrepText = New ArrayList
                            Dim dt As DataTable = pFCE.GetDiscrepText(strWhere).Tables(0)
                            If Not dt Is Nothing Then
                                dsCitations.Tables.Add("DISCREPTEXT")
                                For Each dCol As DataColumn In dt.Columns
                                    dsCitations.Tables("DISCREPTEXT").Columns.Add(dCol.ColumnName, dCol.DataType)
                                Next
                                Dim dr As DataRow
                                For Each dRow As DataRow In dt.Rows
                                    dr = dsCitations.Tables("DISCREPTEXT").NewRow
                                    For Each dCol As DataColumn In dt.Columns
                                        dr(dcol.ColumnName) = dRow.Item(dcol.ColumnName)
                                    Next
                                    dsCitations.Tables("DISCREPTEXT").Rows.Add(dr)
                                    If Not alDiscrepText.Contains(dr("DISCREP TEXT")) Then
                                        alDiscrepText.Add("DISCREP TEXT")
                                    End If
                                Next
                                AddDiscrepTextRelation()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub AddDiscrepTextRelation()
        Try
            Dim dsRel As DataRelation
            dsRel = New DataRelation("CitationToDiscrepText", dsCitations.Tables(0).Columns("CITATION_ID"), dsCitations.Tables(1).Columns("CITATION_ID"), False)
            dsCitations.Relations.Add(dsRel)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub EnableSave(ByVal bolValue As Boolean)
        btnSave.Enabled = bolValue Or bolcitationModified
    End Sub
#End Region

#Region "Citations"
    Private Sub btnCitationDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCitationDelete.Click
        Try
            If Not (ugCitations.ActiveRow Is Nothing) Then
                Dim results As MsgBoxResult
                If ugCitations.ActiveRow.Band.Index = 0 Then
                    results = MsgBox("Do you want to delete citation: " + ugCitations.ActiveRow.Cells("StateCitation").Value.ToString, MsgBoxStyle.YesNo)
                Else
                    results = MsgBox("Do you want to delete discrep: " + ugCitations.ActiveRow.Cells("DISCREP TEXT").Value.ToString, MsgBoxStyle.YesNo)
                End If
                If results = MsgBoxResult.Yes Then
                    bolcitationModified = True
                    If ugCitations.ActiveRow.Band.Index = 0 Then
                        If alCitationPenalty.Contains(ugCitations.ActiveRow.Cells("CITATION_ID").Value) Then
                            alCitationPenalty.Remove(ugCitations.ActiveRow.Cells("CITATION_ID").Value)
                        End If
                        If ugCitations.ActiveRow.Cells("CITATION_ID").Value = 19 Or ugCitations.ActiveRow.Cells("CITATION_ID").Text = "19" Then
                            ' remove all discrep text
                            alDiscrepText = New ArrayList
                        End If
                        bolcitationModified = True
                        EnableSave(True)
                    Else
                        If alDiscrepText.Contains(ugCitations.ActiveRow.Cells("DISCREP TEXT").Text) Then
                            alDiscrepText.Remove(ugCitations.ActiveRow.Cells("DISCREP TEXT").Text)
                        End If
                        bolcitationModified = True
                        EnableSave(True)
                    End If
                    'If nFCEID > 0 Then
                    '    ' save to db
                    '    Dim colCitation As MUSTER.Info.InspectionCitationsCollection
                    '    colCitation = oInspectionCitation.RetrieveByOtherID(nInspectionID, nFCEID, , )
                    '    For Each InsCitationInfo As MUSTER.Info.InspectionCitationInfo In colCitation.Values
                    '        If InsCitationInfo.CitationID = ugCitations.ActiveRow.Cells("CITATION_ID").Value Then
                    '            oInspectionCitation.InspectionCitationInfo = InsCitationInfo
                    '            oInspectionCitation.Deleted = True
                    '            oInspectionCitation.Save()
                    '        End If
                    '    Next
                    'End If
                    ugCitations.ActiveRow.Delete(False)
                End If
            Else
                MsgBox("Select a citaiton / discrep to delete")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCitationAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCitationAdd.Click
        Try
            If pFCE.FacilityID <= 0 Then
                MsgBox("Please select a Facility")
                Exit Sub
            End If
            frmCitation = New CitationList("FCE", pFCE, , alCitationPenalty)
            frmCitation.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCitations_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugCitations.InitializeLayout
        Try
            ugCitations.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
            ugCitations.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugCitations.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False

            ugCitations.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            ugCitations.DisplayLayout.Bands(0).Columns("CITATION_ID").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("FederalCitation").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("Section").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("Small").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("Medium").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("Large").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("CorrectiveAction").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("EPA").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("Created_By").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("DATE_Created").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("LAST_EDITED_By").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Hidden = True
            ugCitations.DisplayLayout.Bands(0).Columns("Deleted").Hidden = True

            ugCitations.DisplayLayout.Bands(0).Columns("StateCitation").Header.Caption = "CITATION"
            ugCitations.DisplayLayout.Bands(0).Columns("Category").Header.Caption = "CATEGORY"
            ugCitations.DisplayLayout.Bands(0).Columns("Description").Header.Caption = "CITATION TEXT"

            ugCitations.DisplayLayout.Bands(0).Columns("StateCitation").Header.VisiblePosition = 0
            ugCitations.DisplayLayout.Bands(0).Columns("Category").Header.VisiblePosition = 1
            ugCitations.DisplayLayout.Bands(0).Columns("Description").Header.VisiblePosition = 2

            If ugCitations.DisplayLayout.Bands.Count > 1 Then
                ugCitations.DisplayLayout.Bands(1).Columns("CITATION_ID").Hidden = True
                ugCitations.DisplayLayout.Bands(1).Columns("QUESTION_ID").Hidden = True

                e.Layout.Bands(1).Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free

                ugCitations.Rows.ExpandAll(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Form Events"
    Private Sub FacilityComplianceEvent_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True
            Populate()
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub FacilityComplianceEvent_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            If pFCE.IsDirty Then
                Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                If Results = MsgBoxResult.Yes Then
                    btnSave.PerformClick()
                ElseIf Results = MsgBoxResult.Cancel Then
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "UI Control Events"
    Private Sub chkSelectedOwner_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectedOwner.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            bolLoading = True
            If chkSelectedOwner.Checked Then
                ' uncheck search for owner
                txtSelectedOwner.Text = strOwnerName

                chkSearchforOwner.Checked = False
                cmbSearchForOwner.DataSource = Nothing
                cmbSearchForOwner.Enabled = False
            Else
                ' check search for owner
                txtSelectedOwner.Text = String.Empty
                chkSearchforOwner.Checked = True
                cmbSearchForOwner.Enabled = True
                PopulateOwners(nOwnerID)
            End If
            PopulateFacility(nOwnerID, nFacID)
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkSearchforOwner_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSearchforOwner.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            bolLoading = True
            If chkSearchforOwner.Checked Then
                ' uncheck selected owner
                txtSelectedOwner.Text = String.Empty
                chkSelectedOwner.Checked = False
                cmbSearchForOwner.Enabled = True
                PopulateOwners(nOwnerID)
            Else
                ' check selected owner
                txtSelectedOwner.Text = strOwnerName

                chkSelectedOwner.Checked = False
                cmbSearchForOwner.DataSource = Nothing
                cmbSearchForOwner.Enabled = False
            End If
            PopulateFacility(nOwnerID, nFacID)
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbSearchForOwner_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSearchForOwner.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pFCE.OwnerID = UIUtilsGen.GetComboBoxValue(cmbSearchForOwner)
            PopulateFacility(pFCE.OwnerID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbFacility_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFacility.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pFCE.FacilityID = UIUtilsGen.GetComboBoxValue(cmbFacility)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtComplianceEventDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtComplianceEventDate.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(dtComplianceEventDate)
            pFCE.FCEDate = UIUtilsGen.GetDatePickerValue(dtComplianceEventDate)
            If Not dtComplianceEventDate.Checked Then
                UIUtilsGen.CreateEmptyFormatDatePicker(dtComplianceEventDate)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            pFCE.Reset()
            bolLoading = True
            alCitationPenalty = New ArrayList
            alDiscrepText = New ArrayList
            Populate()
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim strErr As String
        Dim nInspCitID As Int64 = 0
        Try
            ' check for required fields
            strErr = String.Empty
            If pFCE.OwnerID <= 0 Then
                strErr = "Owner" + vbCrLf
            End If
            If pFCE.FacilityID <= 0 Then
                strErr = "Facility" + vbCrLf
            End If
            If Date.Compare(pFCE.FCEDate, CDate("01/01/0001")) = 0 Then
                strErr = "Compliance Event Date"
            End If
            If ugCitations.Rows.Count <= 0 Then
                strErr += "Citations"
            End If
            If strErr.Length > 0 Then
                strErr = "The following are Required" + vbCrLf + strErr
                MsgBox(strErr)
            Else
                ' if add, inspection id = 0
                oInspection.Retrieve(nInspectionID)

                Dim qIDforCitation19 As Integer = oInspection.CheckListMaster.RetrieveByCheckListItemNum("99999").ID
                Dim qIDforCitationNot19 As Integer = oInspection.CheckListMaster.RetrieveByCheckListItemNum("99998").ID

                If nFCEID = 0 Then

                    oInspection.FacilityID = pFCE.FacilityID
                    oInspection.OwnerID = pFCE.OwnerID
                    oInspection.LetterGenerated = False

                    If oInspection.ID <= 0 Then
                        ' TODO - I do not know why it is 1132. 
                        ' There is noting in the property master of that value, need to check with kevin
                        oInspection.InspectionType = 1132
                    End If

                    If oInspection.ID <= 0 Then
                        oInspection.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oInspection.ModifiedBy = MusterContainer.AppUser.ID
                    End If

                    oInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    pFCE.InspectionID = oInspection.ID
                    If pFCE.ID <= 0 Then
                        oInspection.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oInspection.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    pFCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCitations.Rows
                        oInspectionCitation.Retrieve(oInspection.InspectionInfo, 0)
                        oInspectionCitation.FacilityID = pFCE.FacilityID
                        oInspectionCitation.FCEID = pFCE.ID
                        oInspectionCitation.InspectionID = oInspection.ID
                        oInspectionCitation.CitationID = ugRow.Cells("Citation_ID").Value
                        ' If Category of citation is not discrepancy then use dummy question id (99998) from tblINS_INSPECTION_CHECKLIST_MASTER
                        ' else use dummy question id (99999) from tblINS_INSPECTION_CHECKLIST_MASTER
                        ' This is done to tie citation with discrep text (if any)
                        If oInspectionCitation.CitationID <> 19 Then
                            oInspectionCitation.QuestionID = qIDforCitationNot19
                        Else
                            oInspectionCitation.QuestionID = qIDforCitation19
                        End If

                        If oInspectionCitation.ID <= 0 Then
                            oInspectionCitation.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            oInspectionCitation.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        oInspectionCitation.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                        nInspCitID = oInspectionCitation.ID

                        If Not ugrow.ChildBands Is Nothing Then
                            For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugrow.ChildBands(0).Rows
                                oInspectionDiscrep.Retrieve(oInspection.InspectionInfo, 0)
                                oInspectionDiscrep.Description = ugChildRow.Cells("DISCREP TEXT").Value.ToString
                                oInspectionDiscrep.InspectionID = oInspection.ID
                                ' this is done to tie citation with discrep text
                                oInspectionDiscrep.QuestionID = qIDforCitation19
                                If oInspectionDiscrep.ID <= 0 Then
                                    oInspectionDiscrep.CreatedBy = MusterContainer.AppUser.ID
                                Else
                                    oInspectionDiscrep.ModifiedBy = MusterContainer.AppUser.ID
                                End If
                                oInspectionDiscrep.InspCitID = nInspCitID

                                oInspectionDiscrep.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If
                            Next
                        End If
                    Next
                Else
                    Dim colInspCitations As MUSTER.Info.InspectionCitationsCollection
                    Dim colInspDiscreps As MUSTER.Info.InspectionDiscrepsCollection

                    colInspCitations = oInspectionCitation.RetrieveByOtherID(nInspectionID, nFCEID, , False)
                    colInspDiscreps = oInspectionDiscrep.RetrieveByOtherID(nInspectionID, False)
                    ' if the collection has the row in ultragrid, then remove citation from collection
                    ' if the collection does not have, add citation
                    Dim bolFound As Boolean
                    Dim oInspCitation As MUSTER.Info.InspectionCitationInfo
                    Dim oInspDiscrep As MUSTER.Info.InspectionDiscrepInfo
                    For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCitations.Rows
                        bolFound = False
                        For Each oInspCitation In colInspCitations.Values
                            If oInspCitation.CitationID = ugRow.Cells("CITATION_ID").Value Then
                                nInspCitID = oInspCitation.ID
                                bolFound = True
                                Exit For
                            End If
                        Next
                        If bolFound Then
                            colInspCitations.Remove(oInspCitation.ID)
                        Else
                            'add
                            oInspectionCitation.Retrieve(oInspection.InspectionInfo, 0)
                            oInspectionCitation.FacilityID = pFCE.FacilityID
                            oInspectionCitation.FCEID = pFCE.ID
                            oInspectionCitation.InspectionID = oInspection.ID
                            oInspectionCitation.CitationID = ugRow.Cells("Citation_ID").Value
                            ' If Category of citation is not discrepancy then use dummy question id (99998) from tblINS_INSPECTION_CHECKLIST_MASTER
                            ' else use dummy question id (99999) from tblINS_INSPECTION_CHECKLIST_MASTER
                            ' This is done to tie citation with discrep text (if any)
                            If oInspectionCitation.CitationID <> 19 Then
                                oInspectionCitation.QuestionID = qIDforCitationNot19
                            Else
                                oInspectionCitation.QuestionID = qIDforCitation19
                            End If

                            If oInspectionCitation.ID <= 0 Then
                                oInspectionCitation.CreatedBy = MusterContainer.AppUser.ID
                            Else
                                oInspectionCitation.ModifiedBy = MusterContainer.AppUser.ID
                            End If
                            oInspectionCitation.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If
                            nInspCitID = oInspectionCitation.ID
                        End If

                        ' discrep
                        If Not ugRow.ChildBands Is Nothing Then
                            For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugRow.ChildBands(0).Rows
                                bolFound = False
                                For Each oInspDiscrep In colInspDiscreps.Values
                                    If oInspDiscrep.QuestionID = ugChildRow.Cells("QUESTION_ID").Value Then
                                        bolFound = True
                                        Exit For
                                    End If
                                Next
                                If bolFound Then
                                    colInspDiscreps.Remove(oInspDiscrep.ID)
                                Else
                                    'add
                                    oInspectionDiscrep.Retrieve(oInspection.InspectionInfo, 0)
                                    oInspectionDiscrep.Description = ugChildRow.Cells("DISCREP TEXT").Value.ToString
                                    oInspectionDiscrep.InspectionID = oInspection.ID
                                    ' this is done to tie citation with discrep text
                                    oInspectionDiscrep.QuestionID = qIDforCitation19

                                    If oInspectionDiscrep.ID <= 0 Then
                                        oInspectionDiscrep.CreatedBy = MusterContainer.AppUser.ID
                                    Else
                                        oInspectionDiscrep.ModifiedBy = MusterContainer.AppUser.ID
                                    End If
                                    oInspectionDiscrep.InspCitID = nInspCitID

                                    oInspectionDiscrep.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                                    If Not UIUtilsGen.HasRights(returnVal) Then
                                        Exit Sub
                                    End If
                                End If

                            Next
                        End If
                    Next

                    ' if collection has items, they are not in the grid, delete
                    If colInspCitations.Count > 0 Then
                        For Each oInspCitation In colInspCitations.Values
                            oInspectionCitation.Retrieve(oInspection.InspectionInfo, oInspCitation.ID)
                            oInspectionCitation.Deleted = True
                            If oInspectionCitation.ID <= 0 Then
                                oInspectionCitation.CreatedBy = MusterContainer.AppUser.ID
                            Else
                                oInspectionCitation.ModifiedBy = MusterContainer.AppUser.ID
                            End If
                            oInspectionCitation.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If

                        Next
                    End If

                    ' Discrep
                    If colInspDiscreps.Count > 0 Then
                        For Each oInspDiscrep In colInspDiscreps.Values
                            oInspectionDiscrep.Retrieve(oInspection.InspectionInfo, oInspDiscrep.ID)
                            oInspectionDiscrep.Deleted = True
                            oInspectionDiscrep.ModifiedBy = MusterContainer.AppUser.ID
                            oInspectionDiscrep.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If
                        Next
                    End If

                End If

                MsgBox("FCE Saved Successfully")
                CallingForm.Tag = "1"
                Me.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCitations_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugCitations.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not (ugCitations.ActiveRow Is Nothing) Then
                If ugCitations.ActiveRow.Band.Index = 0 Then
                    If ugCitations.ActiveRow.Cells("Category").Text.ToUpper = "DISCREPANCY" Then
                        ' add discrep text
                        frmCitation = New CitationList("FCE", pFCE, , alDiscrepText, True)
                        frmCitation.ShowDialog()
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "External Events"
    Private Sub pFCE_FacilityComplianceEventChanged(ByVal bolValue As Boolean) Handles pFCE.FacilityComplianceEventChanged
        EnableSave(bolValue)
    End Sub
    Private Sub pFCE_ColChanged(ByVal bolValue As Boolean) Handles pFCE.ColChanged
        EnableSave(bolValue)
    End Sub
    Private Sub frmCitation_evtCitationSelected(ByVal CitationID As Integer, ByVal isDiscrep As Boolean, ByVal strDiscrepText As String) Handles frmCitation.evtCitationSelected
        Try
            If isDiscrep Then
                If Not alDiscrepText.Contains(strDiscrepText) Then
                    alDiscrepText.Add(strDiscrepText)
                End If
            Else
                If Not alCitationPenalty.Contains(CitationID) Then
                    alCitationPenalty.Add(CitationID)
                End If
            End If
            PopulateCitations()
            bolcitationModified = True
            EnableSave(True)
            'If nFCEID <= 0 Then
            '    alCitationPenalty.Add(CitationID)
            '    PopulateCitations()
            '    bolcitationModified = True
            '    EnableSave(True)
            'Else
            '    ' save to db
            '    Dim colCitation As MUSTER.Info.InspectionCitationsCollection
            '    oInspection.Retrieve(nInspectionID)
            '    colCitation = oInspectionCitation.RetrieveByOtherID(nInspectionID, nFCEID, , )
            '    Dim bolFound As Boolean = False
            '    For Each InsCitationInfo As MUSTER.Info.InspectionCitationInfo In colCitation.Values
            '        If InsCitationInfo.CitationID = CitationID Then
            '            bolFound = True
            '        End If
            '    Next
            '    If Not bolFound Then
            '        'oInspection.Retrieve
            '        oInspectionCitation.Retrieve(New MUSTER.Info.InspectionInfo, 0)
            '        oInspectionCitation.FacilityID = pFCE.FacilityID
            '        oInspectionCitation.FCEID = pFCE.ID
            '        oInspectionCitation.InspectionID = oInspection.ID
            '        oInspectionCitation.CitationID = CitationID
            '        oInspectionCitation.Save()
            '    End If
            '    PopulateCitations()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

End Class
