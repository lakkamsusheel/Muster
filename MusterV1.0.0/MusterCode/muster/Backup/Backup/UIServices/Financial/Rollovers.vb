Imports System.IO
Imports Infragistics.Shared
Imports Infragistics.Win
Imports Infragistics.Win.UltraWinGrid
Public Class Rollovers
    Inherits System.Windows.Forms.Form

    ' 1 = Rollover & Zeroout
    ' 2 = New Rollover PO's
    Friend Mode As Int16
    Private nRows As Int16
    Dim returnVal As String = String.Empty
    Protected DOC_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue & "\"
    Dim bolLoading As Boolean = False
    Dim strTitle As String = String.Empty

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents lblCommitments As System.Windows.Forms.Label
    Friend WithEvents pnlRollOversBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlrolloverDetails As System.Windows.Forms.Panel
    Friend WithEvents btnProcess As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents ugCommitments As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlRolloverTop As System.Windows.Forms.Panel
    Friend WithEvents btnPrintGrid As System.Windows.Forms.Button
    Friend WithEvents btnRequest As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlRollOversBottom = New System.Windows.Forms.Panel
        Me.btnPrintGrid = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnProcess = New System.Windows.Forms.Button
        Me.pnlrolloverDetails = New System.Windows.Forms.Panel
        Me.ugCommitments = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.lblCommitments = New System.Windows.Forms.Label
        Me.pnlRolloverTop = New System.Windows.Forms.Panel
        Me.btnRequest = New System.Windows.Forms.Button
        Me.pnlRollOversBottom.SuspendLayout()
        Me.pnlrolloverDetails.SuspendLayout()
        CType(Me.ugCommitments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlRolloverTop.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlRollOversBottom
        '
        Me.pnlRollOversBottom.Controls.Add(Me.btnRequest)
        Me.pnlRollOversBottom.Controls.Add(Me.btnPrintGrid)
        Me.pnlRollOversBottom.Controls.Add(Me.btnPrint)
        Me.pnlRollOversBottom.Controls.Add(Me.btnClose)
        Me.pnlRollOversBottom.Controls.Add(Me.btnProcess)
        Me.pnlRollOversBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlRollOversBottom.Location = New System.Drawing.Point(0, 462)
        Me.pnlRollOversBottom.Name = "pnlRollOversBottom"
        Me.pnlRollOversBottom.Size = New System.Drawing.Size(872, 40)
        Me.pnlRollOversBottom.TabIndex = 2
        '
        'btnPrintGrid
        '
        Me.btnPrintGrid.Location = New System.Drawing.Point(616, 8)
        Me.btnPrintGrid.Name = "btnPrintGrid"
        Me.btnPrintGrid.Size = New System.Drawing.Size(152, 23)
        Me.btnPrintGrid.TabIndex = 6
        Me.btnPrintGrid.Text = "Print Grid"
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(448, 8)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(152, 23)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "Print Report"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(792, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        '
        'btnProcess
        '
        Me.btnProcess.Location = New System.Drawing.Point(16, 8)
        Me.btnProcess.Name = "btnProcess"
        Me.btnProcess.Size = New System.Drawing.Size(200, 23)
        Me.btnProcess.TabIndex = 3
        Me.btnProcess.Text = "Process"
        '
        'pnlrolloverDetails
        '
        Me.pnlrolloverDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlrolloverDetails.Controls.Add(Me.ugCommitments)
        Me.pnlrolloverDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlrolloverDetails.Location = New System.Drawing.Point(0, 32)
        Me.pnlrolloverDetails.Name = "pnlrolloverDetails"
        Me.pnlrolloverDetails.Size = New System.Drawing.Size(872, 430)
        Me.pnlrolloverDetails.TabIndex = 0
        '
        'ugCommitments
        '
        Me.ugCommitments.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCommitments.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugCommitments.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCommitments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCommitments.Location = New System.Drawing.Point(0, 0)
        Me.ugCommitments.Name = "ugCommitments"
        Me.ugCommitments.Size = New System.Drawing.Size(868, 426)
        Me.ugCommitments.TabIndex = 1
        '
        'lblCommitments
        '
        Me.lblCommitments.Location = New System.Drawing.Point(24, 8)
        Me.lblCommitments.Name = "lblCommitments"
        Me.lblCommitments.Size = New System.Drawing.Size(100, 17)
        Me.lblCommitments.TabIndex = 0
        Me.lblCommitments.Text = "Commitments"
        '
        'pnlRolloverTop
        '
        Me.pnlRolloverTop.Controls.Add(Me.lblCommitments)
        Me.pnlRolloverTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlRolloverTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlRolloverTop.Name = "pnlRolloverTop"
        Me.pnlRolloverTop.Size = New System.Drawing.Size(872, 32)
        Me.pnlRolloverTop.TabIndex = 3
        '
        'btnRequest
        '
        Me.btnRequest.Location = New System.Drawing.Point(224, 8)
        Me.btnRequest.Name = "btnRequest"
        Me.btnRequest.Size = New System.Drawing.Size(208, 23)
        Me.btnRequest.TabIndex = 7
        Me.btnRequest.Text = "Request Purchase Order"
        Me.btnRequest.Visible = False
        '
        'Rollovers
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(872, 502)
        Me.Controls.Add(Me.pnlrolloverDetails)
        Me.Controls.Add(Me.pnlRolloverTop)
        Me.Controls.Add(Me.pnlRollOversBottom)
        Me.Name = "Rollovers"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Rollovers"
        Me.pnlRollOversBottom.ResumeLayout(False)
        Me.pnlrolloverDetails.ResumeLayout(False)
        CType(Me.ugCommitments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlRolloverTop.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            ProcessReportString()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot open the file in Word: " + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
        End Try
    End Sub

    Private Sub LoadReportGrid()
        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim dtTotals As DataTable
        Dim oFinancialEvent As New MUSTER.BusinessLogic.pFinancial
        If Mode = 1 Then
            dsLocal = oFinancialEvent.PopulateRolloversZeroesList
            Me.btnProcess.Text = "Process Zero-outs"
        Else
            dsLocal = oFinancialEvent.PopulateRolloverForNewPO
            Me.btnProcess.Text = "Process New PO Numbers"
            Me.btnRequest.Visible = True

        End If

        ugCommitments.DataSource = dsLocal

        ugCommitments.Rows.CollapseAll(True)
        ugCommitments.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Name").Hidden = True
        'ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Number").Hidden = True
        ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_FirstName").Hidden = True
        ugCommitments.DisplayLayout.Override.HeaderClickAction = HeaderClickAction.SortMulti
        ugCommitments.DisplayLayout.Bands(0).Columns("commitmentID").CellActivation = Activation.NoEdit

        If Mode = 1 Then
            ugCommitments.DisplayLayout.Bands(0).Columns("FIN_EVENT_ID").Hidden = True
            ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_ID").Hidden = True
            ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Name").Hidden = False
            ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Number").Hidden = False
            ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Name").CellActivation = Activation.NoEdit
            ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Number").CellActivation = Activation.NoEdit

            ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").CellActivation = Activation.NoEdit
            ugCommitments.DisplayLayout.Bands(0).Columns("Balance").CellActivation = Activation.NoEdit



        Else
            '  ugCommitments.DisplayLayout.Bands(0).Columns("OldPO").CellActivation = Activation.NoEdit
            ugCommitments.DisplayLayout.Bands(0).Columns("OldPO").CellActivation = Activation.Disabled
            ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").CellActivation = Activation.Disabled
            ugCommitments.DisplayLayout.Bands(0).Columns("Activity").CellActivation = Activation.Disabled
            ugCommitments.DisplayLayout.Bands(0).Columns("Balance").CellActivation = Activation.Disabled
            ugCommitments.DisplayLayout.Bands(0).Columns("Reimburse ERAC").CellActivation = Activation.Disabled
            ugCommitments.DisplayLayout.Bands(0).Columns("CommitmentID").CellActivation = Activation.Disabled


            ugCommitments.DisplayLayout.Bands(0).Columns("Rollover").Hidden = True
            ugCommitments.DisplayLayout.Bands(0).Columns("ERACNUM").Hidden = True
            ugCommitments.DisplayLayout.Bands(0).Columns("ERACNAME").Hidden = True
            ugCommitments.DisplayLayout.Bands(0).Columns("ActivityDesc").Hidden = True
            ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Number").Hidden = True
            ugCommitments.DisplayLayout.Bands(0).Columns("facname").Hidden = True


        End If
        ugCommitments.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        '  ugCommitments.DisplayLayout.Bands(0).Columns("Balance").CellActivation = Activation.NoEdit

        If ugCommitments.Rows.Count > 0 Then


            ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Name").Width = 250
            '   ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Number").Width = 110
            ugCommitments.DisplayLayout.Bands(0).Columns("Activity").Width = 140
            ' ugCommitments.DisplayLayout.Bands(0).Columns("Activity").CellActivation = Activation.NoEdit
            '  ugCommitments.DisplayLayout.Bands(0).Columns("Fac_ID").CellActivation = Activation.NoEdit

            'ugTechReports.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            'ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_ID").Hidden = True
            'ugCommitments.DisplayLayout.Bands(0).Columns("Paid").Hidden = True

            'ugTechReports.DisplayLayout.Bands(0).Columns("Received_date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'ugTechReports.DisplayLayout.Bands(0).Columns("Received_date").Header.Caption = "Received"
            'ugTechReports.DisplayLayout.Bands(0).Columns("Received_date").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            'ugTechReports.DisplayLayout.Bands(0).Columns("Received_date").Width = 100
            'ugTechReports.DisplayLayout.Bands(0).Columns("Requested_Amount").Width = 100
            'ugTechReports.DisplayLayout.Bands(0).Columns("Requested_Invoiced").Width = 100
            'ugTechReports.DisplayLayout.Bands(0).Columns("Paid").Width = 100

        End If

    End Sub

    Private Sub Rollovers_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadReportGrid()
    End Sub

    Private Sub ProcessReportString()
        Dim strReturn As String
        Dim strReturn2 As String
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oTechDox As New MUSTER.BusinessLogic.pLustEventDocument
        Dim strTitle As String
        Dim strTitle2 As String
        Dim strData As String
        Dim i As Integer = 0
        Dim oFinancial As New MUSTER.BusinessLogic.pFinancial
        Dim columns As Short
        nRows = 0
        Try
            'Print zero out report ordered by facility ID
            strTitle2 = String.Empty
            ugCommitments.DisplayLayout.Bands(0).SortedColumns.Clear()
            If Mode = 1 Then
                'strTitle = "Rollover & Zero Out Report - Order By Facility"
                strTitle = "Rollover Report - Order By Facility"
                strTitle2 = "Zero Out Report - Order By Facility"
                ugCommitments.DisplayLayout.Bands(0).Columns("RollOver").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("ZeroOut").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ' strReturn = "Fac ID|Vendor#|Vendor|Activity|PO #|Balance|Rollover|ZeroOut"
                strReturn = "Fac ID|Vendor#|Vendor|Activity|PO #|Balance|Rollover"
                strReturn2 = "Fac ID|Vendor#|Vendor|Activity|PO #|Balance|ZeroOut"
                nRows = 1
                columns = 7
            Else    'Mode = 2
                strTitle = "Rollover - New Purchase Order Number Report - Order By Facility"
                ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                strReturn = "Fac ID|Activity|Balance|Old PO#|New PO#|Rollover"
                nRows = 1
                columns = 6
            End If
            For i = 0 To ugCommitments.Rows.Count - 1
                If Mode = 1 Then
                    nRows += 1
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("FAC_ID").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("FAC_ID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Vendor_Number").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("Vendor_Number").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Vendor_Name").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("Vendor_Name").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Activity").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("Activity").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("PONumber").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("PONumber").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Balance").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("Balance").Value)
                    strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("RollOver").Value, "Yes", "No")
                    strReturn2 &= "|" & IIf(ugCommitments.Rows(i).Cells("ZeroOut").Value, "Yes", "No")
                    '  strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("ZeroOut").Value, "Yes", "No")
                Else 'Mode = 2
                    nRows += 1
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("FAC_ID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Activity").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Balance").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("OldPO").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("NewPO").Value)
                    strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("RollOver").Value, "Yes", "No")
                End If
            Next

            Dim oLetter As New Reg_Letters
            oLetter.GenerateGenericLetter(616, strTitle, strReturn, columns, True, , , , , , , Word.WdTableFormat.wdTableFormatContemporary, True, False, True, False, True, False, False, False, True)
            If strTitle2 <> String.Empty Then
                oLetter.GenerateGenericLetter(616, strTitle2, strReturn2, columns, True, , , , , , , Word.WdTableFormat.wdTableFormatContemporary, True, False, True, False, True, False, False, False, True)
            End If

            'Print zero out report ordered by Vendor Name
            strTitle2 = String.Empty
            ugCommitments.DisplayLayout.Bands(0).SortedColumns.Clear()
            If Mode = 1 Then
                '   strTitle = "Rollover & Zero Out Report - Order By Vendor Name"
                strTitle = "Rollover Report - Order By Vendor Name"
                strTitle2 = "Zero Out Report - Order By Vendor Name"
                'strReturn = "Fac ID|Vendor#|Vendor|Activity|PO #|Balance|Rollover|ZeroOut"
                strReturn = "Fac ID|Vendor#|Vendor|Activity|PO #|Balance|Rollover"
                strReturn2 = "Fac ID|Vendor#|Vendor|Activity|PO #|Balance|ZeroOut"
                ugCommitments.DisplayLayout.Bands(0).Columns("RollOver").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("ZeroOut").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Name").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Number").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled

                For i = 0 To ugCommitments.Rows.Count - 1
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("FAC_ID").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("FAC_ID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Vendor_Number").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("Vendor_Number").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Vendor_Name").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("Vendor_Name").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Activity").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("Activity").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("PONumber").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("PONumber").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Balance").Value)
                    strReturn2 &= "|" & CStr(ugCommitments.Rows(i).Cells("Balance").Value)
                    strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("RollOver").Value, "Yes", "No")
                    strReturn2 &= "|" & IIf(ugCommitments.Rows(i).Cells("ZeroOut").Value, "Yes", "No")
                    ' strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("ZeroOut").Value, "Yes", "No")
                Next
            Else 'Mode = 2 
                strTitle = "Rollover - New Purchase Order Number Report - Order By PONumber"
                strReturn = "Fac ID|Activity|Balance|Old PO#|New PO#|Rollover"
                ugCommitments.DisplayLayout.Bands(0).Columns("OldPO").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                For i = 0 To ugCommitments.Rows.Count - 1
                    'strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("CommitmentID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("FAC_ID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Activity").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Balance").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("OldPO").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("NewPO").Value)
                    strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("RollOver").Value, "Yes", "No")
                Next
            End If
            oLetter.GenerateGenericLetter(616, strTitle, strReturn, columns, True, , , , , , , Word.WdTableFormat.wdTableFormatContemporary, True, False, True, False, True, False, False, False, True)
            If strTitle2 <> String.Empty Then
                oLetter.GenerateGenericLetter(616, strTitle2, strReturn2, columns, True, , , , , , , Word.WdTableFormat.wdTableFormatContemporary, True, False, True, False, True, False, False, False, True)
            End If
            ugCommitments.DisplayLayout.Bands(0).SortedColumns.Clear()

            If Mode = 1 Then
                ugCommitments.DisplayLayout.Bands(0).Columns("RollOver").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                ugCommitments.DisplayLayout.Bands(0).Columns("ZeroOut").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            Else    'Mode = 2
                ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            End If


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Print Data " + ex.Message, ex))
            MyErr.ShowDialog()
        Finally

        End Try
    End Sub
    Private Sub btnProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcess.Click

        If MsgBox("Have you printed out the grid or reports? If not, this command will not proceed.", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            If Mode = 1 Then
                ProcessRolloverZeroOut()
            Else
                ProcessRolloverNewPO()
            End If

            LoadReportGrid()
        End If

    End Sub

    Private Sub btnRequest_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRequest.Click

        If Mode = 2 Then
            Me.ProcessPoRequest()
        End If

    End Sub

    Private Sub ProcessRolloverZeroOut()
        Dim oLetter As New Reg_Letters
        Dim strTitle As String
        Dim strData As String
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oTechDox As New MUSTER.BusinessLogic.pLustEventDocument
        Dim oCommitment As New MUSTER.BusinessLogic.pFinancialCommitment
        Dim oAdjustment As New MUSTER.BusinessLogic.pFinancialCommitAdjustment
        Dim oFinancial As New MUSTER.BusinessLogic.pFinancial
        Dim oTechEvent As New MUSTER.BusinessLogic.pLustEvent
        Dim oFacility As New MUSTER.BusinessLogic.pFacility
        Dim oOwner As New MUSTER.BusinessLogic.pOwner
        Dim file As File
        Dim dtTable As New DataTable
        Dim dr As DataRow
        Dim filePath As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Templates).ProfileValue & "\Financial\UnencumberanceMemoTemplate.doc"
        Try

            strData = ""
            dtTable.Columns.Add("FacilityID")
            dtTable.Columns.Add("TecEventID")
            dtTable.Columns.Add("FinEventID")
            dtTable.Columns.Add("CommitmentID")
            dtTable.Columns.Add("CommitAdjustID")
            ugCommitments.DisplayLayout.Bands(0).SortedColumns.Clear()
            ugCommitments.DisplayLayout.Bands(0).Columns("Rollover").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            If Not System.IO.Directory.Exists(DOC_PATH) Then
                MsgBox("Path does not exists to generate letter")
                Exit Sub
            End If

            If file.Exists(filePath) Then
                For Each ugrow In ugCommitments.Rows
                    Dim nFinEventID As Integer = 0
                    Dim pConstruct As New MUSTER.BusinessLogic.pContactStruct
                    Dim strVendorName As String = String.Empty
                    If ugrow.Cells("Rollover").Value = True Or ugrow.Cells("Rollover").Text = "True" Then

                        oCommitment.Retrieve(ugrow.Cells("CommitmentID").Value)
                        oCommitment.RollOver = True
                        If oCommitment.CommitmentID <= 0 Then
                            oCommitment.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            oCommitment.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        oCommitment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                        'strData &= "|" & CStr(ugrow.Cells("CommitmentID").Value)
                        strData &= "|" & CStr(ugrow.Cells("FAC_ID").Value)
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'If Not (ugrow.Cells("FIN_EVENT_ID").Value Is System.DBNull.Value) Then
                        '    nFinEventID = Integer.Parse(ugrow.Cells("FIN_EVENT_ID").Value)
                        'End If
                        'If nFinEventID > 0 Then
                        '    Dim drRow As DataRow
                        '    oFinancial.Retrieve(nFinEventID)
                        '    Dim dtVendor As DataTable = pConstruct.GETContactName(nFinEventID, 32, 616)  'oLetter.GetXHAndXLContacts(nFinEventID, 32, 616) 'oVendor.GetByID(oFinancialEvent.VendorID, nFinEventID, 616)
                        '    If dtVendor.Rows.Count > 0 Then
                        '        For Each drRow In dtVendor.Rows
                        '            If drRow.Item("Type") = 1185 Then
                        '                If Not (drRow.Item("AssocCompany") Is System.DBNull.Value Or drRow.Item("AssocCompany") = String.Empty) Then
                        '                    If drRow.Item("IsPerson") = True Then
                        '                        'strContactName = drRow.Item("CONTACT_Name")
                        '                    End If
                        '                    strVendorName = drRow.Item("AssocCompany")

                        '                Else
                        '                    strVendorName = drRow.Item("CONTACT_Name") 'IIf(drRow.Item("Title") = String.Empty, "", drRow.Item("Title") + " ") + drRow.Item("First_Name") + " " + IIf(drRow.Item("Middle_Name") = String.Empty, "", drRow.Item("Middle_Name") + " ") + drRow.Item("Last_Name") + IIf(drRow.Item("Suffix") = String.Empty, "", " " + drRow.Item("Suffix"))
                        '                End If
                        '            End If
                        '        Next
                        '    End If
                        'End If
                        strData &= "|" & CStr(ugrow.Cells("Vendor_Number").Value)
                        strData &= "|" & CStr(ugrow.Cells("Vendor_Name").Value)
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        strData &= "|" & CStr(ugrow.Cells("Activity").Value)
                        strData &= "|" & CStr(ugrow.Cells("PONumber").Value)
                        strData &= "|" & CStr(ugrow.Cells("Balance").Value)
                        'strData &= "|" & IIf(ugrow.Cells("RollOver").Value, "Yes", "No")
                        'strData &= "|" & IIf(ugrow.Cells("ZeroOut").Value, "Yes", "No")
                    End If
                    If ugrow.Cells("ZeroOut").Value = True Or ugrow.Cells("ZeroOut").Text = "True" Then

                        oAdjustment.Retrieve(0)
                        oAdjustment.CommitmentID = ugrow.Cells("CommitmentID").Value
                        oAdjustment.DirectorApprovalReq = False
                        oAdjustment.FinancialApprovalReq = False
                        oAdjustment.AdjustDate = Now.Date
                        oAdjustment.AdjustType = 1076 'Unencumberance
                        oAdjustment.AdjustAmount = ugrow.Cells("Balance").Value
                        oAdjustment.Comments = "Zero Out"
                        If oAdjustment.CommitmentID <= 0 Then
                            oAdjustment.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            oAdjustment.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        oAdjustment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)

                        oCommitment.Retrieve(oAdjustment.CommitmentID)
                        oFinancial.Retrieve(oCommitment.Fin_Event_ID)
                        oTechEvent.Retrieve(oFinancial.TecEventID)
                        oFacility.Retrieve(oOwner.OwnerInfo, oTechEvent.FacilityID, "SELF", "FACILITY")
                        oOwner.Retrieve(oFacility.OwnerID)
                        dr = dtTable.NewRow
                        dr("FacilityID") = oTechEvent.FacilityID
                        dr("TecEventID") = oFinancial.TecEventID
                        dr("FinEventID") = oCommitment.Fin_Event_ID
                        dr("CommitmentID") = oAdjustment.CommitmentID
                        dr("CommitAdjustID") = oAdjustment.CommitAdjustmentID
                        dtTable.Rows.Add(dr)
                        'oLetter.GenerateFinancialLetter(oTechEvent.FacilityID, "Change Order Unencumberance Memo", "UnencumberanceMemo_" & oAdjustment.CommitAdjustmentID, "Change Order Unencumberance Memo", "UnencumberanceMemoTemplate.doc", oOwner, oFinancial.TecEventID, oCommitment.Fin_Event_ID, oAdjustment.CommitmentID, 0, oAdjustment.CommitAdjustmentID, 0)
                    End If
                    'If (ugrow.Cells("ZeroOut").Value = False And ugrow.Cells("Rollover").Value = False) Then
                    '    oCommitment.Retrieve(ugrow.Cells("CommitmentID").Value)
                    '    If oCommitment.RollOver = True Then
                    '        oCommitment.RollOver = False
                    '        If oCommitment.CommitmentID <= 0 Then
                    '            oCommitment.CreatedBy = MusterContainer.AppUser.ID
                    '        Else
                    '            oCommitment.ModifiedBy = MusterContainer.AppUser.ID
                    '        End If
                    '        oCommitment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    '        If Not UIUtilsGen.HasRights(returnVal) Then
                    '            Exit Sub
                    '        End If

                    '    End If
                    'End If
                Next
                If strData <> String.Empty Then
                    strTitle = "Rollover Report - Order By Facility"
                    'strData = "ID|Fac ID|Vendor|Activity|PO #|Balance|Rollover|ZeroOut" & strData
                    strData = "Fac ID|Vendor #|Vendor|Activity|PO #|Balance" & strData
                    oLetter.GenerateFinancialGenericLetter(UIUtilsGen.ModuleID.Financial, strTitle, strData, 6, False, , , , , "Rollover Report - Order By Facility", "Rollover Report", Word.WdTableFormat.wdTableFormatContemporary, True, False, True, False, True, False, False, False, True)
                End If


                'Process Report 2 (sorted by Vendor Name)
                strData = ""
                ugCommitments.DisplayLayout.Bands(0).SortedColumns.Clear()
                ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("Rollover").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Name").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ugCommitments.DisplayLayout.Bands(0).Columns("Vendor_Number").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled

                For Each ugrow In ugCommitments.Rows
                    Dim nFinEventID As Integer = 0
                    Dim pConstruct As New MUSTER.BusinessLogic.pContactStruct
                    Dim strVendorName As String = String.Empty
                    Dim strSortName As String = String.Empty
                    If ugrow.Cells("Rollover").Value = True Or ugrow.Cells("Rollover").Text = "True" Then
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        strData &= "|" & CStr(ugrow.Cells("FAC_ID").Value)
                        strData &= "|" & CStr(ugrow.Cells("Vendor_Number").Value)
                        strData &= "|" & CStr(ugrow.Cells("Vendor_Name").Value)
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        strData &= "|" & CStr(ugrow.Cells("Activity").Value)
                        strData &= "|" & CStr(ugrow.Cells("PONumber").Value)
                        strData &= "|" & CStr(ugrow.Cells("Balance").Value)
                        'strData &= "|" & IIf(ugrow.Cells("RollOver").Value, "Yes", "No")
                        'strData &= "|" & IIf(ugrow.Cells("ZeroOut").Value, "Yes", "No")
                    End If
                Next

                'For Each ugrow In ugCommitments.Rows
                '    Dim nFinEventID As Integer = 0
                '    Dim pConstruct As New MUSTER.BusinessLogic.pContactStruct
                '    Dim strVendorName As String = String.Empty
                '    If ugrow.Cells("Rollover").Value = True Then
                '        'strData &= "|" & CStr(ugrow.Cells("CommitmentID").Value)
                '        strData &= "|" & CStr(ugrow.Cells("FAC_ID").Value)
                '        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '        If Not (ugrow.Cells("FIN_EVENT_ID").Value Is System.DBNull.Value) Then
                '            nFinEventID = Integer.Parse(ugrow.Cells("FIN_EVENT_ID").Value)
                '        End If
                '        If nFinEventID > 0 Then
                '            Dim drRow As DataRow
                '            oFinancial.Retrieve(nFinEventID)
                '            Dim dtVendor As DataTable = pConstruct.GETContactName(nFinEventID, 32, 616)  'oLetter.GetXHAndXLContacts(nFinEventID, 32, 616) 'oVendor.GetByID(oFinancialEvent.VendorID, nFinEventID, 616)
                '            If dtVendor.Rows.Count > 0 Then
                '                For Each drRow In dtVendor.Rows
                '                    If drRow.Item("Type") = 1185 Then
                '                        If Not (drRow.Item("AssocCompany") Is System.DBNull.Value Or drRow.Item("AssocCompany") = String.Empty) Then
                '                            If drRow.Item("IsPerson") = True Then
                '                                'strContactName = drRow.Item("CONTACT_Name")
                '                            End If
                '                            strVendorName = drRow.Item("AssocCompany")

                '                        Else
                '                            strVendorName = drRow.Item("CONTACT_Name") 'IIf(drRow.Item("Title") = String.Empty, "", drRow.Item("Title") + " ") + drRow.Item("First_Name") + " " + IIf(drRow.Item("Middle_Name") = String.Empty, "", drRow.Item("Middle_Name") + " ") + drRow.Item("Last_Name") + IIf(drRow.Item("Suffix") = String.Empty, "", " " + drRow.Item("Suffix"))
                '                        End If
                '                    End If
                '                Next
                '            End If
                '        End If
                '        strData &= "|" & strVendorName
                '        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '        strData &= "|" & CStr(ugrow.Cells("Activity").Value)
                '        strData &= "|" & CStr(ugrow.Cells("PONumber").Value)
                '        strData &= "|" & CStr(ugrow.Cells("Balance").Value)
                '        'strData &= "|" & IIf(ugrow.Cells("RollOver").Value, "Yes", "No")
                '        'strData &= "|" & IIf(ugrow.Cells("ZeroOut").Value, "Yes", "No")
                '    End If
                'Next




                If dtTable.Rows.Count > 0 Then
                    'oLetter.GenerateFinancialLetter(oTechEvent.FacilityID, "Change Order Unencumberance Memo", "UnencumberanceMemo_" & oAdjustment.CommitAdjustmentID, "Change Order Unencumberance Memo", "UnencumberanceMemoTemplate.doc", oOwner, oFinancial.TecEventID, oCommitment.Fin_Event_ID, oAdjustment.CommitmentID, 0, oAdjustment.CommitAdjustmentID, 0)
                    oLetter.GenerateFinancialUnencumberanceMemoTemplate(dtTable, "Change Order Unencumberance Memo", "UnencumberanceMemo_" & oAdjustment.CommitAdjustmentID, "Change Order Unencumberance Memo", "UnencumberanceMemoTemplate.doc", oOwner)
                End If
                If strData <> String.Empty Then
                    strTitle = "Rollover Report - Order By Vendor"
                    'strData = "ID|Fac ID|Vendor|Activity|PO #|Balance|Rollover|ZeroOut" & strData
                    strData = "Fac ID|Vendor #|Vendor|Activity|PO #|Balance" & strData
                    oLetter.GenerateFinancialGenericLetter(UIUtilsGen.ModuleID.Financial, strTitle, strData, 6, False, , , , , "Rollover Report - Order By Vendor", "Rollover Report", Word.WdTableFormat.wdTableFormatContemporary, True, False, True, False, True, False, False, False, True)
                End If
            Else
                MsgBox("File does not exists to generate letter")
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Process Rollover/ZeroOut " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub



    Private Sub ProcessPoRequest()
        Dim oLetter As New Reg_Letters
        Dim strTitle As String
        Dim strData As String
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        Dim file As File
        Dim dtTable As New DataTable
        Dim dr As DataRow
        Dim filePath As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Templates).ProfileValue & "\Financial\CommitmentMemoTemplate.doc"
        Try



            If Not System.IO.Directory.Exists(DOC_PATH) Then
                MsgBox("Path does not exists to generate letter")
                Exit Sub
            End If



            dtTable.Columns.Add("Fac_ID")
            dtTable.Columns.Add("facName")
            dtTable.Columns.Add("Balance")
            dtTable.Columns.Add("Vendor_Name")
            dtTable.Columns.Add("Vendor_Number")
            dtTable.Columns.Add("ERACNUM")
            dtTable.Columns.Add("ERACNAME")
            dtTable.Columns.Add("REIMBURSE ERAC")
            dtTable.Columns.Add("CommitmentID")
            dtTable.Columns.Add("OldPO")
            dtTable.Columns.Add("ID")

            If file.Exists(filePath) Then
                Dim cnt As Integer = 0
                For Each ugrow In ugCommitments.Rows
                    Dim nFinEventID As Integer = 0

                    dr = dtTable.NewRow
                    For Each column As DataColumn In dtTable.Columns
                        If column.ColumnName <> "ID" Then
                            dr.Item(column.ColumnName) = ugrow.Cells(column.ColumnName).Value
                        End If
                    Next
                    dr.Item("ID") = String.Format("{0}{1}", "0".PadLeft(10, "0").Substring(0, 10 - cnt.ToString.Length), cnt)

                    dtTable.Rows.Add(dr)

                    cnt += 1
                Next

                If dtTable.Rows.Count > 0 Then
                    'oLetter.GenerateFinancialLetter(oTechEvent.FacilityID, "Change Order Unencumberance Memo", "UnencumberanceMemo_" & oAdjustment.CommitAdjustmentID, "Change Order Unencumberance Memo", "UnencumberanceMemoTemplate.doc", oOwner, oFinancial.TecEventID, oCommitment.Fin_Event_ID, oAdjustment.CommitmentID, 0, oAdjustment.CommitAdjustmentID, 0)
                    oLetter.GenerateFinancialPoRequest(dtTable, "Purchase Order Request", "PoRequestMemo_" & String.Format("{0:d}", Date.Now).Replace("/", "_").Replace(" ", "_"), "New PO Request Memo", "CommitmentMemoTemplate.doc")
                End If
            Else
                MsgBox("File does not exists to generate letter")
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Process Order Change. " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub



    Private Sub ProcessRolloverNewPO()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oCommitment As New MUSTER.BusinessLogic.pFinancialCommitment
        Dim oAdjustment As New MUSTER.BusinessLogic.pFinancialCommitAdjustment
        Dim sRolloverOffset As Single
        Dim strOldPO As Single
        Try

            'If Not System.IO.Directory.Exists(DOC_PATH) Then
            '    MsgBox("Path does not exists to generate letter")
            '    Exit Sub
            'End If
            For Each ugrow In ugCommitments.Rows
                If ugrow.Cells("Rollover").Value = True Then
                    If ugrow.Cells("NewPO").Value > String.Empty Then
                        oCommitment.Retrieve(ugrow.Cells("CommitmentID").Value)

                        oCommitment.PONumber = ugrow.Cells("NewPO").Text
                        oCommitment.RollOver = False

                        'oAdjustment.Retrieve(0)
                        'oAdjustment.CommitmentID = oCommitment.CommitmentID
                        'oAdjustment.AdjustAmount = ugrow.Cells("Balance").Value
                        'oAdjustment.AdjustType = 1076 'Unemcumberance
                        'oAdjustment.Comments = "Offset for Rollover to PO:  " & CStr(ugrow.Cells("NewPO").Value)
                        'oAdjustment.AdjustDate = Now.Date

                        'sRolloverOffset = oCommitment.GetCommitmentOffset(ugrow.Cells("CommitmentID").Value)
                        'strOldPO = oCommitment.PONumber

                        'oAdjustment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        'If Not UIUtilsGen.HasRights(returnVal) Then
                        '    Exit Sub
                        'End If

                        'oCommitment.PONumber = ugrow.Cells("NewPO").Value
                        'oCommitment.RollOver = False
                        'oCommitment.RollOverID()

                        If oCommitment.CommitmentID <= 0 Then
                            oCommitment.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            oCommitment.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        oCommitment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        'If sRolloverOffset <> 0 Then
                        '    oAdjustment.Retrieve(0)
                        '    oAdjustment.CommitmentID = oCommitment.CommitmentID
                        '    If sRolloverOffset < 0 Then
                        '        sRolloverOffset = sRolloverOffset * -1
                        '        oAdjustment.AdjustType = 1076 'Unemcumberance
                        '    Else
                        '        oAdjustment.AdjustType = 1075
                        '        oAdjustment.Approved = True
                        '    End If
                        '    oAdjustment.AdjustAmount = sRolloverOffset
                        '    oAdjustment.Comments = "Offset for Rollover from PO:  " + CStr(strOldPO)
                        '    oAdjustment.AdjustDate = Now.Date
                        '    oAdjustment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        '    If Not UIUtilsGen.HasRights(returnVal) Then
                        '        Exit Sub
                        '    End If
                        'End If
                    End If
                Else
                    oCommitment.Retrieve(ugrow.Cells("CommitmentID").Value)
                    If oCommitment.RollOver = True Then
                        oCommitment.RollOver = False
                        If oCommitment.CommitmentID <= 0 Then
                            oCommitment.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            oCommitment.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        oCommitment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Process Rollover - New PO " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub ugCommitments_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugCommitments.CellChange
        If Mode = 1 Then
            If bolLoading Then Exit Sub
            Dim bolLoadingLocal As Boolean = bolLoading
            Try
                bolLoading = True
                If "ZEROOUT".Equals(e.Cell.Column.Key.ToUpper) Then
                    If e.Cell.Text.ToUpper = "TRUE" And Not e.Cell.Value Then
                        e.Cell.Row.Cells("ROLLOVER").Value = False
                    End If
                    e.Cell.Value = e.Cell.Text
                ElseIf "ROLLOVER".Equals(e.Cell.Column.Key.ToUpper) Then
                    If e.Cell.Text.ToUpper = "TRUE" And Not e.Cell.Value Then
                        e.Cell.Row.Cells("ZEROOUT").Value = False
                    End If
                    e.Cell.Value = e.Cell.Text
                End If
                bolLoading = False

            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
                bolLoading = False
            End Try


        End If


        If Mode = 2 Then
            Try
                If "REIMBURSE ERAC".Equals(e.Cell.Column.Key.ToUpper) Then
                    If e.Cell.Row.Cells("ERACNAME").Value.ToString.Length = 0 Then
                        e.Cell.Value = 0
                    End If

                End If

            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
            End Try
        End If
    End Sub
    Private Sub ProcessReportString1()
        Dim strReturn As String
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oTechDox As New MUSTER.BusinessLogic.pLustEventDocument
        Dim strTitle As String
        Dim strData As String
        Dim i As Integer = 0
        nRows = 0
        Try
            If Mode = 1 Then
                strTitle = "Rollover & Zero Out Report"
                ugCommitments.DisplayLayout.Bands(0).Columns("ZeroOut").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                strReturn = "ID|Fac ID|Activity|PO #|Balance|Rollover|ZeroOut"
                nRows = 1
            Else
                strTitle = "Rollover - New Purchase Order Number Report - Order By Facility"
                ugCommitments.DisplayLayout.Bands(0).Columns("FAC_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                strReturn = "ID|Fac ID|Activity|Balance|Old PO#|New PO#|Rollover"
                nRows = 1
            End If

            For i = 0 To ugCommitments.Rows.Count - 1
                If Mode = 1 Then
                    'If ugrow.Cells("Rollover").Value = True Or ugrow.Cells("ZeroOut").Value = True Then
                    nRows += 1
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("CommitmentID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("FAC_ID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Activity").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("PONumber").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Balance").Value)
                    strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("RollOver").Value, "Yes", "No")
                    strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("ZeroOut").Value, "Yes", "No")
                    'End If
                End If

                If Mode = 2 Then
                    'If ugrow.Cells("Rollover").Value = True Then 'Or ugrow.Cells("ZeroOut").Value = True Then
                    nRows += 1
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("CommitmentID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("FAC_ID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Activity").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Balance").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("OldPO").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("NewPO").Value)
                    strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("RollOver").Value, "Yes", "No")
                    'End If
                End If
            Next

            Dim oLetter As New Reg_Letters
            oLetter.GenerateGenericLetter(616, strTitle, strReturn, 7, True, , , , , , , Word.WdTableFormat.wdTableFormatContemporary, True, False, True, False, True, False, False, False, True)

            If Mode = 2 Then
                strTitle = "Rollover - New Purchase Order Number Report - Order By PONumber"
                strReturn = "ID|Fac ID|Activity|Balance|Old PO#|New PO#|Rollover"
                ugCommitments.DisplayLayout.Bands(0).Columns("OldPO").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                For i = 0 To ugCommitments.Rows.Count - 1
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("CommitmentID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("FAC_ID").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Activity").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("Balance").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("OldPO").Value)
                    strReturn &= "|" & CStr(ugCommitments.Rows(i).Cells("NewPO").Value)
                    strReturn &= "|" & IIf(ugCommitments.Rows(i).Cells("RollOver").Value, "Yes", "No")
                Next
                oLetter.GenerateGenericLetter(616, strTitle, strReturn, 7, True, , , , , , , Word.WdTableFormat.wdTableFormatContemporary, True, False, True, False, True, False, False, False, True)
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Print Data " + ex.Message, ex))
            MyErr.ShowDialog()
        Finally

        End Try
    End Sub

    Private Class MySortComparer
        Implements IComparer

        Public Sub New()
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare

            ' Passed in objects are cells. So you have to typecast them to UltraGridCell objects first.
            Dim xCell As UltraGridCell = DirectCast(x, UltraGridCell)
            Dim yCell As UltraGridCell = DirectCast(y, UltraGridCell)

            ' Do your own comparision between the values of xCell and yCell and return a negative
            ' number if xCell is less than yCell, positive number if xCell is greater than yCell,
            ' and 0 if xCell and yCell are equal.

            ' Following code does an case-insensitive compare of the values converted to string.
            Dim text1 As String = xCell.Value.ToString()
            Dim text2 As String = yCell.Value.ToString()

            Return String.Compare(text1, text2, True)

        End Function

    End Class

    Private Sub btnPrintGrid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintGrid.Click
        Me.ugCommitments.PrintPreview()

    End Sub

    Private Sub ugCommitments_InitializePrintPreview(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelablePrintPreviewEventArgs) Handles ugCommitments.InitializePrintPreview
        ' Set the zomm level to 100 % in the print preview.
        If Mode = 1 Then
            e.PrintPreviewSettings.Zoom = 1.3
            strTitle = "Rollover And ZeroOut Report"
        Else
            e.PrintPreviewSettings.Zoom = 1.2
            strTitle = "Rollover New Purchase Order No Report"
        End If

        ' Set the location and size of the print preview dialog.
        e.PrintPreviewSettings.DialogLeft = SystemInformation.WorkingArea.X
        e.PrintPreviewSettings.DialogTop = SystemInformation.WorkingArea.Y
        e.PrintPreviewSettings.DialogWidth = SystemInformation.WorkingArea.Width
        e.PrintPreviewSettings.DialogHeight = SystemInformation.WorkingArea.Height

        ' Horizontally fit everything in a signle page.
        e.DefaultLogicalPageLayoutInfo.FitWidthToPages = 1

        ' Set up the header and the footer.
        e.DefaultLogicalPageLayoutInfo.PageHeader = strTitle
        e.DefaultLogicalPageLayoutInfo.PageHeaderHeight = 40
        e.DefaultLogicalPageLayoutInfo.PageHeaderAppearance.FontData.SizeInPoints = 14
        e.DefaultLogicalPageLayoutInfo.PageHeaderAppearance.TextHAlign = HAlign.Center
        e.DefaultLogicalPageLayoutInfo.PageHeaderBorderStyle = UIElementBorderStyle.Solid

        ' Use <#> token in the string to designate page numbers.
        e.DefaultLogicalPageLayoutInfo.PageFooter = "Page <#>."
        e.DefaultLogicalPageLayoutInfo.PageFooterHeight = 40
        e.DefaultLogicalPageLayoutInfo.PageFooterAppearance.TextHAlign = HAlign.Right
        e.DefaultLogicalPageLayoutInfo.PageFooterAppearance.FontData.Italic = DefaultableBoolean.True
        e.DefaultLogicalPageLayoutInfo.PageFooterBorderStyle = UIElementBorderStyle.Solid

        ' Set the ClippingOverride to Yes.
        e.DefaultLogicalPageLayoutInfo.ClippingOverride = ClippingOverride.Yes

        ' Set the document name through the PrintDocument which returns a PrintDocument object.
        e.PrintDocument.DocumentName = strTitle

    End Sub

End Class




