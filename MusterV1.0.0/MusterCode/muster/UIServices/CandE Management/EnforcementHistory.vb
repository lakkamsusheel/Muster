Public Class EnforcementHistory
    Inherits System.Windows.Forms.Form

#Region "User Declared Variables"
    Private bolIsFromOCE As Boolean = True
    Private bolLoading As Boolean = True

    ' Variables for OCE Enforcement History
    Private pOCE As MUSTER.BusinessLogic.pOwnerComplianceEvent
    Private nOwnerID As Integer
    Private strOwnerName As String
    Private nLCELicenseeID As Integer
    ' Variables for LCE Enforcement History
    Private pLCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal isFromOCE As Boolean, Optional ByVal ownerID As Integer = 0, Optional ByVal strOwnName As String = "", Optional ByVal oce As MUSTER.BusinessLogic.pOwnerComplianceEvent = Nothing, Optional ByVal LCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent = Nothing, Optional ByVal LCELicenseeID As Integer = 0)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        bolLoading = True
        bolIsFromOCE = isFromOCE
        If bolIsFromOCE Then
            If oce Is Nothing Then
                pOCE = New MUSTER.BusinessLogic.pOwnerComplianceEvent
            Else
                pOCE = oce
            End If
            nOwnerID = ownerID
            strOwnerName = strOwnName
        Else
            If LCE Is Nothing Then
                pLCE = New MUSTER.BusinessLogic.pLicenseeComplianceEvent
            Else
                pLCE = LCE
            End If
            nLCELicenseeID = LCELicenseeID
        End If
        bolLoading = False
    End Sub

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
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents btnExpandAll As System.Windows.Forms.Button
    Friend WithEvents ugEnforcementHistory As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.btnExpandAll = New System.Windows.Forms.Button
        Me.ugEnforcementHistory = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTop.SuspendLayout()
        CType(Me.ugEnforcementHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.btnExpandAll)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(792, 40)
        Me.pnlTop.TabIndex = 0
        '
        'btnExpandAll
        '
        Me.btnExpandAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExpandAll.Location = New System.Drawing.Point(704, 8)
        Me.btnExpandAll.Name = "btnExpandAll"
        Me.btnExpandAll.Size = New System.Drawing.Size(80, 23)
        Me.btnExpandAll.TabIndex = 0
        Me.btnExpandAll.Text = "Expand All"
        '
        'ugEnforcementHistory
        '
        Me.ugEnforcementHistory.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugEnforcementHistory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugEnforcementHistory.Location = New System.Drawing.Point(0, 40)
        Me.ugEnforcementHistory.Name = "ugEnforcementHistory"
        Me.ugEnforcementHistory.Size = New System.Drawing.Size(792, 533)
        Me.ugEnforcementHistory.TabIndex = 1
        '
        'EnforcementHistory
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(792, 573)
        Me.Controls.Add(Me.ugEnforcementHistory)
        Me.Controls.Add(Me.pnlTop)
        Me.MinimizeBox = False
        Me.Name = "EnforcementHistory"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Enforcement History"
        Me.pnlTop.ResumeLayout(False)
        CType(Me.ugEnforcementHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "UI Support Routines"
    Private Sub ExpandAll(ByVal bol As Boolean, ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid, ByRef btn As Button)
        If bol Then
            btn.Text = "Collapse All"
            ug.Rows.ExpandAll(True)
        Else
            btn.Text = "Expand All"
            ug.Rows.CollapseAll(True)
        End If
    End Sub
    Private Sub Populate()
        Try
            If bolIsFromOCE Then
                ugEnforcementHistory.DataSource = pOCE.GetPriorEnforcements(nOwnerID, False)
                ExpandAll(False, ugEnforcementHistory, btnExpandAll)
                If nOwnerID = 0 Then
                    Me.Text += " for All Owners"
                Else
                    Me.Text += " for Owner : " + strOwnerName + " (" + nOwnerID.ToString + ")"
                End If
            Else

                pLCE.GetEnforcementHistory(nLCELicenseeID)
                ugEnforcementHistory.DataSource = pLCE.EntityTable(False, True)

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "UI Control Events"
    Private Sub btnExpandAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpandAll.Click
        Try
            If btnExpandAll.Text = "Expand All" Then
                ExpandAll(True, ugEnforcementHistory, btnExpandAll)
            Else
                ExpandAll(False, ugEnforcementHistory, btnExpandAll)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugEnforcementHistory_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugEnforcementHistory.InitializeLayout
        If bolLoading Then Exit Sub
        Try
            If bolIsFromOCE Then
                e.Layout.UseFixedHeaders = True
                e.Layout.Override.FixedHeaderIndicator = Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None
                'e.Layout.Bands(0).Columns("SELECTED").Header.Fixed = True
                e.Layout.Bands(0).Columns("STATUS").Header.Fixed = True

                ' cannot have fixed header if row header is split into two rows
                'e.Layout.Bands(1).Columns("SELECTED").Header.Fixed = True
                'e.Layout.Bands(1).Columns("OWNERNAME").Header.Fixed = True
                'e.Layout.Bands(1).Columns("ENSITE ID").Header.Fixed = True
                'e.Layout.Bands(1).Columns("FILLER1").Header.Fixed = True

                e.Layout.Bands(2).Columns("SELECTED").Header.Fixed = True
                e.Layout.Bands(2).Columns("OWNER_ID").Header.Fixed = True
                e.Layout.Bands(2).Columns("FACILITY_ID").Header.Fixed = True
                e.Layout.Bands(2).Columns("FACILITY").Header.Fixed = True

                e.Layout.Bands(0).Override.RowAppearance.BackColor = Color.White
                e.Layout.Bands(1).Override.RowAppearance.BackColor = Color.RosyBrown
                e.Layout.Bands(1).Override.RowAlternateAppearance.BackColor = Color.PeachPuff
                e.Layout.Bands(2).Override.RowAppearance.BackColor = Color.Khaki

                e.Layout.Bands(1).Columns("OWNERNAME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                e.Layout.Bands(1).Columns("OCE DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                e.Layout.Bands(2).Columns("FACILITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                e.Layout.Bands(0).Columns("OCE_STATUS").Hidden = True
                e.Layout.Bands(0).Columns("SELECTED").Hidden = True

                e.Layout.Bands(1).Columns("SELECTED").Hidden = True
                e.Layout.Bands(1).Columns("OCE_STATUS").Hidden = True
                e.Layout.Bands(1).Columns("OWNER_ID").Hidden = True
                e.Layout.Bands(1).Columns("OCE_ID").Hidden = True

                e.Layout.Bands(2).Columns("SELECTED").Hidden = True
                e.Layout.Bands(2).Columns("OCE_ID").Hidden = True
                e.Layout.Bands(2).Columns("INS_CIT_ID").Hidden = True

                e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

                'e.Layout.Bands(1).Columns("RESCINDED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                e.Layout.Bands(1).Columns("OCE DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                e.Layout.Bands(1).Columns("NEXT DUE DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                e.Layout.Bands(1).Columns("LAST PROCESS DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                e.Layout.Bands(1).Columns("DATE RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                'e.Layout.Bands(2).Columns("RESCINDED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                e.Layout.Bands(2).Columns("DUE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
                e.Layout.Bands(2).Columns("RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit

                e.Layout.Bands(1).Columns("FILLER1").Header.Caption = ""
                'e.Layout.Bands(1).Columns("FILLER2").Header.Caption = ""
                e.Layout.Bands(1).Columns("FILLER3").Header.Caption = ""
                e.Layout.Bands(1).Columns("FILLER4").Header.Caption = ""
                e.Layout.Bands(1).Columns("FILLER5").Header.Caption = ""

                e.Layout.Bands(1).Columns("OCE DATE").Header.Caption = "OCE" + vbCrLf + "DATE"
                e.Layout.Bands(1).Columns("LAST PROCESS DATE").Header.Caption = "LAST" + vbCrLf + "PROCESS DATE"
                e.Layout.Bands(1).Columns("NEXT DUE DATE").Header.Caption = "NEXT" + vbCrLf + "DUE DATE"
                e.Layout.Bands(1).Columns("OVERRIDE DUE DATE").Header.Caption = "OVERRIDE" + vbCrLf + "DUE DATE"
                e.Layout.Bands(1).Columns("POLICY AMOUNT").Header.Caption = "POLICY" + vbCrLf + "AMOUNT"
                e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").Header.Caption = "OVERRIDE" + vbCrLf + "AMOUNT"
                e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").Header.Caption = "SETTLEMENT" + vbCrLf + "AMOUNT"

                e.Layout.Bands(1).Columns("PAID AMOUNT").Header.Caption = "PAID" + vbCrLf + "AMOUNT"
                e.Layout.Bands(1).Columns("DATE RECEIVED").Header.Caption = "DATE" + vbCrLf + "RECEIVED"
                e.Layout.Bands(1).Columns("WORKSHOP DATE").Header.Caption = "WORKSHOP" + vbCrLf + "DATE"
                e.Layout.Bands(1).Columns("WORKSHOP RESULT").Header.Caption = "WORKSHOP" + vbCrLf + "RESULT"
                e.Layout.Bands(1).Columns("SHOW CAUSE HEARING DATE").Header.Caption = "SHOW CAUSE" + vbCrLf + "HEARING DATE"
                e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").Header.Caption = "SHOW CAUSE" + vbCrLf + "HEARING RESULT"
                e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").Header.Caption = "COMMISSION" + vbCrLf + "HEARING RESULT"
                e.Layout.Bands(1).Columns("COMMISSION HEARING DATE").Header.Caption = "COMMISSION" + vbCrLf + "HEARING DATE"
                e.Layout.Bands(1).Columns("AGREED ORDER #").Header.Caption = "AGREED" + vbCrLf + "ORDER #"
                e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").Header.Caption = "ADMINISTRATIVE" + vbCrLf + "ORDER #"
                e.Layout.Bands(1).Columns("PENDING LETTER").Header.Caption = "PENDING" + vbCrLf + "LETTER"
                e.Layout.Bands(1).Columns("LETTER PRINTED").Header.Caption = "LETTER" + vbCrLf + "PRINTED"
                e.Layout.Bands(1).Columns("LETTER GENERATED").Header.Caption = "LETTER" + vbCrLf + "GENERATED"

                e.Layout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
                e.Layout.Bands(1).Override.DefaultColWidth = 100
                e.Layout.Bands(1).ColHeaderLines = 2
                e.Layout.Bands(1).Override.RowSelectorWidth = 1

                'e.Layout.Bands(1).Columns("PENDING LETTER").CellMultiLine = Infragistics.Win.DefaultableBoolean.True
                'e.Layout.Bands(1).Columns("PENDING LETTER").RowLayoutColumnInfo.SpanY = 2
                e.Layout.Bands(1).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree

                If e.Layout.Bands(1).Groups.Count = 0 Then
                    e.Layout.Bands(1).Groups.Add("ROW1")
                    'e.layout.Bands(1).Groups.Add("ROW2")
                End If

                e.Layout.Bands(1).GroupHeadersVisible = False

                'e.Layout.Bands(1).Columns("SELECTED").Group = e.Layout.Bands(1).Groups("ROW1")
                'e.Layout.Bands(1).Columns("FILLER1").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("OWNERNAME").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("ENSITE ID").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("COMMENTS").Group = e.Layout.Bands(1).Groups("ROW1")
                'e.Layout.Bands(1).Columns("FILLER2").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("RESCINDED").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("FILLER3").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("OCE DATE").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("LAST PROCESS DATE").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("NEXT DUE DATE").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("OVERRIDE DUE DATE").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("STATUS").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("ESCALATION").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("POLICY AMOUNT").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("FILLER4").Group = e.Layout.Bands(1).Groups("ROW1")

                e.Layout.Bands(1).Columns("PAID AMOUNT").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("DATE RECEIVED").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("WORKSHOP DATE").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("WORKSHOP RESULT").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("SHOW CAUSE HEARING DATE").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("COMMISSION HEARING DATE").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("AGREED ORDER #").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("PENDING LETTER").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("FILLER5").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("LETTER GENERATED").Group = e.Layout.Bands(1).Groups("ROW1")
                e.Layout.Bands(1).Columns("LETTER PRINTED").Group = e.Layout.Bands(1).Groups("ROW1")

                e.Layout.Bands(1).LevelCount = 3
                'e.Layout.Bands(1).Columns("SELECTED").Level = 0
                'e.Layout.Bands(1).Columns("FILLER1").Level = 1
                e.Layout.Bands(1).Columns("OWNERNAME").Level = 0
                e.Layout.Bands(1).Columns("ENSITE ID").Level = 1
                e.Layout.Bands(1).Columns("COMMENTS").Level = 2
                'e.Layout.Bands(1).Columns("FILLER2").Level = 1
                e.Layout.Bands(1).Columns("RESCINDED").Level = 0
                e.Layout.Bands(1).Columns("FILLER3").Level = 1
                e.Layout.Bands(1).Columns("OCE DATE").Level = 0
                e.Layout.Bands(1).Columns("LAST PROCESS DATE").Level = 1
                e.Layout.Bands(1).Columns("NEXT DUE DATE").Level = 0
                e.Layout.Bands(1).Columns("OVERRIDE DUE DATE").Level = 1
                e.Layout.Bands(1).Columns("STATUS").Level = 0
                e.Layout.Bands(1).Columns("ESCALATION").Level = 1
                e.Layout.Bands(1).Columns("POLICY AMOUNT").Level = 0
                e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").Level = 1
                e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").Level = 0
                e.Layout.Bands(1).Columns("FILLER4").Level = 1
                e.Layout.Bands(1).Columns("PAID AMOUNT").Level = 0
                e.Layout.Bands(1).Columns("DATE RECEIVED").Level = 1
                e.Layout.Bands(1).Columns("WORKSHOP DATE").Level = 0
                e.Layout.Bands(1).Columns("WORKSHOP RESULT").Level = 1
                e.Layout.Bands(1).Columns("SHOW CAUSE HEARING DATE").Level = 0
                e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").Level = 1
                e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").Level = 1
                e.Layout.Bands(1).Columns("COMMISSION HEARING DATE").Level = 0
                e.Layout.Bands(1).Columns("AGREED ORDER #").Level = 0
                e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").Level = 1
                e.Layout.Bands(1).Columns("PENDING LETTER").Level = 0
                e.Layout.Bands(1).Columns("FILLER5").Level = 1
                e.Layout.Bands(1).Columns("LETTER GENERATED").Level = 0
                e.Layout.Bands(1).Columns("LETTER PRINTED").Level = 1

                e.Layout.Bands(3).Columns("INS_CIT_ID").Hidden = True
                e.Layout.Bands(3).Columns("INSPECTION_ID").Hidden = True
                e.Layout.Bands(3).Columns("CITATION_ID").Hidden = True
                e.Layout.Bands(3).Columns("QUESTION_ID").Hidden = True
                e.Layout.Bands(3).Columns("INS_DESCREP_ID").Hidden = True
                e.Layout.Bands(3).Columns("DISCREP TEXT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                e.Layout.Bands(3).Columns("DISCREP TEXT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                e.Layout.Bands(3).Columns("DISCREP TEXT").Width = 400
                e.Layout.Bands(3).Columns("RESCINDED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                e.Layout.Bands(3).Columns("RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            Else
                InitializeLayout(ugEnforcementHistory)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Form Events"
    Private Sub EnforcementHistory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Populate()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub InitializeLayout(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        ' Dim col As UltraGridColumn
        Try

            'ugGrid.DisplayLayout.Bands(0).Columns.Add("--")
            'ugGrid.DisplayLayout.Bands(0).Columns.Add("Status")

            If ugGrid.DisplayLayout.Bands(1).Groups.Count = 0 Then
                ugGrid.DisplayLayout.Bands(1).Groups.Add("Row1")
                ugGrid.DisplayLayout.Bands(1).Groups.Add("Row2")
            End If

            'After you have bound your grid to a DataSource you should create an unbound column that will be used as your CheckBox column. In the InitializeLayout event add the following code to create an unbound column:
            ugGrid.DisplayLayout.Bands(0).Columns.Add("Selected").DataType = GetType(Boolean)
            ugGrid.DisplayLayout.Bands(0).Columns("Selected").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
            ugGrid.DisplayLayout.Bands(0).Columns("Selected").Header.VisiblePosition = 0

            ugGrid.DisplayLayout.Bands(1).Columns.Add("Selected").DataType = GetType(Boolean)
            ugGrid.DisplayLayout.Bands(1).Columns("Selected").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
            ugGrid.DisplayLayout.Bands(1).Columns("Selected").Header.VisiblePosition = 0

            ugGrid.DisplayLayout.Bands(1).Override.RowAppearance.BackColor = Color.Yellow
            ugGrid.DisplayLayout.Bands(2).Override.RowAppearance.BackColor = Color.Khaki
            'Me.ugGrid.SupportThemes = True
            'For Each col In Me.ugGrid.DisplayLayout.Bands(0).Columns
            '    col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent
            '    col.Header.Appearance.BackColor = Color.DarkGray
            'Next
            'For Each col In Me.ugGrid.DisplayLayout.Bands(1).Columns
            '    col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent
            '    col.Header.Appearance.BackColor = Color.DarkGray
            'Next
            'For Each col In Me.ugGrid.DisplayLayout.Bands(2).Columns
            '    col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent
            '    col.Header.Appearance.BackColor = Color.DarkGray
            'Next

            ugGrid.DisplayLayout.Bands(1).Columns("Selected").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER5").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Licensee").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER1").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Rescinded").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER2").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("LCE" + vbCrLf + "Date").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Last" + vbCrLf + "Process Date").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Next" + vbCrLf + "Due Date").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Override" + vbCrLf + "Due Date").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("LCE_Status").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Escalation").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Policy" + vbCrLf + "Amount").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Override" + vbCrLf + "Amount").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER3").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Paid" + vbCrLf + "Amount").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("Date" + vbCrLf + "Received").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row1")
            ugGrid.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")
            ugGrid.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")
            ugGrid.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Date").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")
            ugGrid.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")
            ugGrid.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Date").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")
            ugGrid.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")
            ugGrid.DisplayLayout.Bands(1).Columns("Pending" + vbCrLf + "Letter").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")
            ugGrid.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Printed").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")
            ugGrid.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Generated").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER4").Group = ugGrid.DisplayLayout.Bands(1).Groups("Row2")

            ugGrid.DisplayLayout.Bands(1).LevelCount = 2
            ugGrid.DisplayLayout.Bands(1).Columns("Selected").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER5").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Licensee").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER1").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Rescinded").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER2").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("LCE" + vbCrLf + "Date").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("Last" + vbCrLf + "Process Date").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Next" + vbCrLf + "Due Date").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("Override" + vbCrLf + "Due Date").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("LCE_Status").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("Escalation").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Policy" + vbCrLf + "Amount").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("Override" + vbCrLf + "Amount").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER3").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Paid" + vbCrLf + "Amount").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("Date" + vbCrLf + "Received").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Date").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Date").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("Pending" + vbCrLf + "Letter").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Printed").Level = 1
            ugGrid.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Generated").Level = 0
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER4").Level = 1

            ugGrid.DisplayLayout.Bands(1).Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            ugGrid.DisplayLayout.Bands(1).Override.CellMultiLine = Infragistics.Win.DefaultableBoolean.True
            ugGrid.DisplayLayout.Bands(1).Override.DefaultColWidth = 80
            ugGrid.DisplayLayout.Bands(1).ColHeaderLines = 2
            ugGrid.DisplayLayout.Bands(1).Override.RowSelectorWidth = 2
            ugGrid.DisplayLayout.Bands(1).Columns("Licensee").CellMultiLine = Infragistics.Win.DefaultableBoolean.False

            ugGrid.DisplayLayout.Bands(1).Columns("FILLER1").Header.Caption = ""
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER1").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER1").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

            ugGrid.DisplayLayout.Bands(1).Columns("FILLER2").Header.Caption = ""
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER2").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER2").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

            ugGrid.DisplayLayout.Bands(1).Columns("FILLER3").Header.Caption = ""
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER3").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER3").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow
            ugGrid.DisplayLayout.Bands(1).GroupHeadersVisible = False

            ugGrid.DisplayLayout.Bands(1).Columns("FILLER4").Header.Caption = ""
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER4").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER4").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

            ugGrid.DisplayLayout.Bands(1).Columns("FILLER5").Header.Caption = ""
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER5").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER5").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

            ugGrid.DisplayLayout.Bands(1).Columns("LCE_ID").Hidden = True
            ugGrid.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").Hidden = True
            ugGrid.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").Hidden = True

            'For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In ugGrid.Rows
            '    If Not dr.ChildBands Is Nothing Then

            '        ' Loop throgh each of the child bands.
            '        Dim childBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand = Nothing
            '        For Each childBand In dr.ChildBands
            '            For Each dr1 As Infragistics.Win.UltraWinGrid.UltraGridRow In childBand.Rows
            '                If dr1.Cells("LCE" + vbCrLf + "Date").Text = "01/01/0001" Then
            '                    dr1.Cells("LCE" + vbCrLf + "Date").Value = DBNull.Value
            '                End If
            '            Next
            '        Next
            '    End If
            'Next

            ugGrid.DisplayLayout.Bands(0).Columns("LCE_Status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            ugGrid.DisplayLayout.Bands(1).Columns("LCE_Status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("Licensee").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("Rescinded").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("LCE" + vbCrLf + "DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("Last" + vbCrLf + "Process Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("Next" + vbCrLf + "Due Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("Escalation").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("Policy" + vbCrLf + "Amount").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'ugGrid.DisplayLayout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("Pending" + vbCrLf + "Letter").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Generated").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'ugGrid.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Printed").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            ugGrid.DisplayLayout.Bands(2).Columns("FACILITY_ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(2).Columns("Facility").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(2).Columns("Citation").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugGrid.DisplayLayout.Bands(2).Columns("Citation Text").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            ugGrid.DisplayLayout.Bands(2).Columns("Citation Text").Width = 400
            ugGrid.DisplayLayout.Bands(2).Columns("Facility").Width = 200
            ugGrid.DisplayLayout.Bands(1).Columns("Licensee").Width = 150
            ugGrid.DisplayLayout.Bands(1).Columns("FILLER1").Width = 150

            ' Set the Style to DropDownList.
            ugGrid.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            ugGrid.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList

            '' populate columns
            'If ugGrid.DisplayLayout.ValueLists.All.Length = 0 Then
            '    ugGrid.DisplayLayout.ValueLists.Add("ATTENDED")
            '    ugGrid.DisplayLayout.ValueLists.Add("NO SHOW")
            'End If

            ' populate the whole column as the table is the same for each row
            ' Show cause hearing results -property type id is 129 
            If ugGrid.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").ValueList Is Nothing Then
                Dim vListShowCauseHearingResult As New Infragistics.Win.ValueList
                For Each row As DataRow In pLCE.getDropDownValues(129).Rows
                    vListShowCauseHearingResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                ugGrid.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").ValueList = vListShowCauseHearingResult
            End If

            If ugGrid.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").ValueList Is Nothing Then
                Dim vListShowCauseHearingResult As New Infragistics.Win.ValueList
                For Each row As DataRow In pLCE.getDropDownValues(130).Rows
                    vListShowCauseHearingResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                ugGrid.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").ValueList = vListShowCauseHearingResult
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

End Class
