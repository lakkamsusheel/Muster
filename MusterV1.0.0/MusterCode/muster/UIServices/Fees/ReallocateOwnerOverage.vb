Public Class ReallocateOwnerOverage
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Friend InvoiceID As Int64
    Friend OwnerID As Int64
    Friend FacilityID As Int64

    Dim oInvoice As New MUSTER.BusinessLogic.pFeeInvoice
    Dim returnVal As String = String.Empty
#End Region
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
    Friend WithEvents pnlReallocateOwnerTop As System.Windows.Forms.Panel
    Friend WithEvents pnlReallocateOwnerBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlReallocateOwnerDetails As System.Windows.Forms.Panel
    Friend WithEvents btnReallocateOverage As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblReallocateCaption As System.Windows.Forms.Label
    Friend WithEvents ugReallocateOverage As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblCurrentOverageValue As System.Windows.Forms.Label
    Friend WithEvents lblCurrentOverage As System.Windows.Forms.Label
    Friend WithEvents lblNewOverageValue As System.Windows.Forms.Label
    Friend WithEvents lblNewOverage As System.Windows.Forms.Label
    Friend WithEvents lblReallocateAmountValue As System.Windows.Forms.Label
    Friend WithEvents lblReallocateAmount As System.Windows.Forms.Label
    Friend WithEvents lblOriginalAmtValue As System.Windows.Forms.Label
    Friend WithEvents lblOriginalAmt As System.Windows.Forms.Label
    Friend WithEvents lblOrigInvoiceValue As System.Windows.Forms.Label
    Friend WithEvents lblOrigInvoice As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlReallocateOwnerTop = New System.Windows.Forms.Panel
        Me.lblCurrentOverageValue = New System.Windows.Forms.Label
        Me.lblCurrentOverage = New System.Windows.Forms.Label
        Me.lblNewOverageValue = New System.Windows.Forms.Label
        Me.lblNewOverage = New System.Windows.Forms.Label
        Me.lblReallocateAmountValue = New System.Windows.Forms.Label
        Me.lblReallocateAmount = New System.Windows.Forms.Label
        Me.lblOriginalAmtValue = New System.Windows.Forms.Label
        Me.lblOriginalAmt = New System.Windows.Forms.Label
        Me.lblOrigInvoiceValue = New System.Windows.Forms.Label
        Me.lblOrigInvoice = New System.Windows.Forms.Label
        Me.lblReallocateCaption = New System.Windows.Forms.Label
        Me.pnlReallocateOwnerBottom = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnReallocateOverage = New System.Windows.Forms.Button
        Me.pnlReallocateOwnerDetails = New System.Windows.Forms.Panel
        Me.ugReallocateOverage = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlReallocateOwnerTop.SuspendLayout()
        Me.pnlReallocateOwnerBottom.SuspendLayout()
        Me.pnlReallocateOwnerDetails.SuspendLayout()
        CType(Me.ugReallocateOverage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlReallocateOwnerTop
        '
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblCurrentOverageValue)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblCurrentOverage)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblNewOverageValue)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblNewOverage)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblReallocateAmountValue)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblReallocateAmount)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblOriginalAmtValue)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblOriginalAmt)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblOrigInvoiceValue)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblOrigInvoice)
        Me.pnlReallocateOwnerTop.Controls.Add(Me.lblReallocateCaption)
        Me.pnlReallocateOwnerTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlReallocateOwnerTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlReallocateOwnerTop.Name = "pnlReallocateOwnerTop"
        Me.pnlReallocateOwnerTop.Size = New System.Drawing.Size(792, 80)
        Me.pnlReallocateOwnerTop.TabIndex = 0
        '
        'lblCurrentOverageValue
        '
        Me.lblCurrentOverageValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCurrentOverageValue.Location = New System.Drawing.Point(136, 40)
        Me.lblCurrentOverageValue.Name = "lblCurrentOverageValue"
        Me.lblCurrentOverageValue.TabIndex = 20
        Me.lblCurrentOverageValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCurrentOverage
        '
        Me.lblCurrentOverage.Location = New System.Drawing.Point(40, 40)
        Me.lblCurrentOverage.Name = "lblCurrentOverage"
        Me.lblCurrentOverage.Size = New System.Drawing.Size(112, 23)
        Me.lblCurrentOverage.TabIndex = 19
        Me.lblCurrentOverage.Text = "Current Overage"
        '
        'lblNewOverageValue
        '
        Me.lblNewOverageValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNewOverageValue.Location = New System.Drawing.Point(640, 40)
        Me.lblNewOverageValue.Name = "lblNewOverageValue"
        Me.lblNewOverageValue.TabIndex = 18
        Me.lblNewOverageValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblNewOverage
        '
        Me.lblNewOverage.Location = New System.Drawing.Point(544, 40)
        Me.lblNewOverage.Name = "lblNewOverage"
        Me.lblNewOverage.Size = New System.Drawing.Size(112, 23)
        Me.lblNewOverage.TabIndex = 17
        Me.lblNewOverage.Text = "Overage Balance"
        '
        'lblReallocateAmountValue
        '
        Me.lblReallocateAmountValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblReallocateAmountValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReallocateAmountValue.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.lblReallocateAmountValue.Location = New System.Drawing.Point(384, 40)
        Me.lblReallocateAmountValue.Name = "lblReallocateAmountValue"
        Me.lblReallocateAmountValue.TabIndex = 16
        Me.lblReallocateAmountValue.Text = "0.00"
        Me.lblReallocateAmountValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblReallocateAmount
        '
        Me.lblReallocateAmount.Location = New System.Drawing.Point(280, 40)
        Me.lblReallocateAmount.Name = "lblReallocateAmount"
        Me.lblReallocateAmount.Size = New System.Drawing.Size(112, 23)
        Me.lblReallocateAmount.TabIndex = 15
        Me.lblReallocateAmount.Text = "Amount Reallocated"
        '
        'lblOriginalAmtValue
        '
        Me.lblOriginalAmtValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOriginalAmtValue.Location = New System.Drawing.Point(568, 72)
        Me.lblOriginalAmtValue.Name = "lblOriginalAmtValue"
        Me.lblOriginalAmtValue.TabIndex = 14
        Me.lblOriginalAmtValue.Text = "0.00"
        Me.lblOriginalAmtValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblOriginalAmtValue.Visible = False
        '
        'lblOriginalAmt
        '
        Me.lblOriginalAmt.Location = New System.Drawing.Point(464, 72)
        Me.lblOriginalAmt.Name = "lblOriginalAmt"
        Me.lblOriginalAmt.TabIndex = 13
        Me.lblOriginalAmt.Text = "Original Amount"
        Me.lblOriginalAmt.Visible = False
        '
        'lblOrigInvoiceValue
        '
        Me.lblOrigInvoiceValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOrigInvoiceValue.Location = New System.Drawing.Point(352, 72)
        Me.lblOrigInvoiceValue.Name = "lblOrigInvoiceValue"
        Me.lblOrigInvoiceValue.TabIndex = 12
        Me.lblOrigInvoiceValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblOrigInvoiceValue.Visible = False
        '
        'lblOrigInvoice
        '
        Me.lblOrigInvoice.Location = New System.Drawing.Point(240, 72)
        Me.lblOrigInvoice.Name = "lblOrigInvoice"
        Me.lblOrigInvoice.TabIndex = 11
        Me.lblOrigInvoice.Text = "Invoice"
        Me.lblOrigInvoice.Visible = False
        '
        'lblReallocateCaption
        '
        Me.lblReallocateCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReallocateCaption.Location = New System.Drawing.Point(312, 8)
        Me.lblReallocateCaption.Name = "lblReallocateCaption"
        Me.lblReallocateCaption.Size = New System.Drawing.Size(168, 23)
        Me.lblReallocateCaption.TabIndex = 0
        Me.lblReallocateCaption.Text = "Reallocate Owner Overage"
        '
        'pnlReallocateOwnerBottom
        '
        Me.pnlReallocateOwnerBottom.Controls.Add(Me.btnCancel)
        Me.pnlReallocateOwnerBottom.Controls.Add(Me.btnReallocateOverage)
        Me.pnlReallocateOwnerBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlReallocateOwnerBottom.Location = New System.Drawing.Point(0, 526)
        Me.pnlReallocateOwnerBottom.Name = "pnlReallocateOwnerBottom"
        Me.pnlReallocateOwnerBottom.Size = New System.Drawing.Size(792, 40)
        Me.pnlReallocateOwnerBottom.TabIndex = 2
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(400, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(120, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        '
        'btnReallocateOverage
        '
        Me.btnReallocateOverage.Location = New System.Drawing.Point(272, 8)
        Me.btnReallocateOverage.Name = "btnReallocateOverage"
        Me.btnReallocateOverage.Size = New System.Drawing.Size(120, 23)
        Me.btnReallocateOverage.TabIndex = 3
        Me.btnReallocateOverage.Text = "Reallocate Overage"
        '
        'pnlReallocateOwnerDetails
        '
        Me.pnlReallocateOwnerDetails.Controls.Add(Me.ugReallocateOverage)
        Me.pnlReallocateOwnerDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlReallocateOwnerDetails.Location = New System.Drawing.Point(0, 80)
        Me.pnlReallocateOwnerDetails.Name = "pnlReallocateOwnerDetails"
        Me.pnlReallocateOwnerDetails.Size = New System.Drawing.Size(792, 446)
        Me.pnlReallocateOwnerDetails.TabIndex = 0
        '
        'ugReallocateOverage
        '
        Me.ugReallocateOverage.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugReallocateOverage.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugReallocateOverage.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugReallocateOverage.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugReallocateOverage.Location = New System.Drawing.Point(0, 0)
        Me.ugReallocateOverage.Name = "ugReallocateOverage"
        Me.ugReallocateOverage.Size = New System.Drawing.Size(792, 446)
        Me.ugReallocateOverage.TabIndex = 1
        '
        'ReallocateOwnerOverage
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(792, 566)
        Me.Controls.Add(Me.pnlReallocateOwnerDetails)
        Me.Controls.Add(Me.pnlReallocateOwnerBottom)
        Me.Controls.Add(Me.pnlReallocateOwnerTop)
        Me.Name = "ReallocateOwnerOverage"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Reallocate Owner Overage"
        Me.pnlReallocateOwnerTop.ResumeLayout(False)
        Me.pnlReallocateOwnerBottom.ResumeLayout(False)
        Me.pnlReallocateOwnerDetails.ResumeLayout(False)
        CType(Me.ugReallocateOverage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Support Routines"

#End Region
#Region "UI Control Events"
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region


    Private Sub ReallocateOwnerOverage_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        oInvoice.Retrieve(InvoiceID)

        lblOrigInvoiceValue.Text = oInvoice.WarrantNumber
        lblOriginalAmtValue.Text = FormatNumber(oInvoice.InvoiceAmount, 2, TriState.True, TriState.False, TriState.True)


        lblCurrentOverageValue.Text = FormatNumber(oInvoice.GetOverpaymentBucket(OwnerID), 2, TriState.True, TriState.False, TriState.True)
        lblNewOverageValue.Text = FormatNumber(lblCurrentOverageValue.Text - lblReallocateAmountValue.Text, 2, TriState.True, TriState.False, TriState.True)

        LoadugReallocateOverage()
    End Sub

    Private Sub LoadugReallocateOverage()
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim dsLocal As DataSet
        Dim tmpBand As Int16

        Try


            dsLocal = oFeeInvoice.GetInvoiceLineItemBalanceDue_ByInvoiceID(InvoiceID)
            dsLocal.Tables(0).Columns("Facility_ID").ReadOnly = True
            dsLocal.Tables(0).Columns("FacilityName").ReadOnly = True
            dsLocal.Tables(0).Columns("Fiscal_Year").ReadOnly = True
            dsLocal.Tables(0).Columns("Charges").ReadOnly = True
            dsLocal.Tables(0).Columns("Balance").ReadOnly = True

            ugReallocateOverage.DataSource = dsLocal

            ugReallocateOverage.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            'ugReallocateOverage.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            'ugReallocateOverage.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("ITEM_SEQ_NUMBER").Hidden = True

                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Facility_ID").TabStop = False

                ugReallocateOverage.DisplayLayout.Bands(0).Columns("FacilityName").TabStop = False
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Fiscal_Year").TabStop = False
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Charges").TabStop = False
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Balance").TabStop = False

                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Charges").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Charges").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center


                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Balance").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Reallocation").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Reallocation").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center


                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Facility_ID").Width = 50
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("FacilityName").Width = 170
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Fiscal_Year").Width = 40
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Charges").Width = 65
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Balance").Width = 65
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Reallocation").Width = 75
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("CheckNumber").Width = 85
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Reason").Width = 175


                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Facility_ID").Header.Caption = "Facility"
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("FacilityName").Header.Caption = "Facility Name"
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("Fiscal_Year").Header.Caption = "FY"
                ugReallocateOverage.DisplayLayout.Bands(0).Columns("CheckNumber").Header.Caption = "Check Number"

            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load ugReallocateOverage " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugReallocateOverage_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugReallocateOverage.AfterCellUpdate
        Dim sTotal As Single
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try

            If ugReallocateOverage.Rows.Count > 0 Then
                For Each ugrow In ugReallocateOverage.Rows
                    If ugrow.Cells("Reallocation").Value > 0 Then
                        If IsNumeric(ugrow.Cells("Reallocation").Value) Then
                            'ugrow.Cells("Credits").Value = FormatNumber(ugrow.Cells("Credits").Value, 2, TriState.True, TriState.False, TriState.True)
                            If ugrow.Cells("Reallocation").Value <= ugrow.Cells("Balance").Value Then
                                sTotal = sTotal + ugrow.Cells("Reallocation").Value
                            Else
                                MsgBox("Reallocation Amount Cannot Be Greater Than Balance.", MsgBoxStyle.OKOnly, "Invalid Amount")
                                ugrow.Cells("Reallocation").Value = 0
                            End If

                        End If
                    End If
                Next
            End If
            lblReallocateAmountValue.Text = FormatNumber(sTotal, 2, TriState.True, TriState.False, TriState.True)
            lblNewOverageValue.Text = FormatNumber(lblCurrentOverageValue.Text - lblReallocateAmountValue.Text, 2, TriState.True, TriState.False, TriState.True)
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot ugReallocateOverage_AfterCellUpdate " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnReallocateOverage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReallocateOverage.Click
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oAdjustment As New MUSTER.BusinessLogic.pFeeAdjustment
        Dim bolReallocation As Boolean

        Try

            If ValidateData() = False Then
                Exit Sub
            End If

            bolReallocation = False

            For Each ugrow In ugReallocateOverage.Rows
                If ugrow.Cells("Reallocation").Value > 0 Then
                    bolReallocation = True

                    oAdjustment.Retrieve(0)

                    oAdjustment.Deleted = False
                    oAdjustment.OwnerID = OwnerID
                    oAdjustment.CreditCode = "AP"
                    oAdjustment.FiscalYear = ugrow.Cells("Fiscal_Year").Value
                    oAdjustment.FacilityID = ugrow.Cells("Facility_ID").Value
                    oAdjustment.InvoiceNumber = oInvoice.WarrantNumber
                    oAdjustment.ItemSeqNumber = ugrow.Cells("ITEM_SEQ_NUMBER").Value
                    oAdjustment.Amount = ugrow.Cells("Reallocation").Value
                    oAdjustment.Applied = Now.Date
                    oAdjustment.CheckNumber = ugrow.Cells("CheckNumber").Value
                    oAdjustment.Reason = ugrow.Cells("Reason").Value

                    oAdjustment.CreatedBy = MusterContainer.AppUser.ID
                    oAdjustment.Save(CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                End If
            Next

            If bolReallocation Then
                MsgBox("Funds have been reallocated.", MsgBoxStyle.OKOnly, "Overage Reallocation Confirmation")
                Me.Close()
            Else
                MsgBox("No reallocation was indicated.  No funds have been reallocated.", MsgBoxStyle.OKOnly, "No Reallocation Records")
            End If


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot btnReallocateOverage_Click " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Function ValidateData() As Boolean
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim bolAnyReason As Boolean
        Dim strLastReason As String

        Try

            If ugReallocateOverage.Rows.Count <= 0 Then
                MsgBox("There is no outstanding balance for the selected invoice.  No funds may be reallocated.", MsgBoxStyle.OKOnly, "Invalid Invoice")
                Return False
            End If

            If lblNewOverageValue.Text < 0 Then
                MsgBox("Reallocation amount exceeds overage.  Please reduce the amount reallocated.", MsgBoxStyle.OKOnly, "Excessive Reallocation")
                Return False
            End If
            bolAnyReason = False

            For Each ugrow In ugReallocateOverage.Rows
                If ugrow.Cells("Reallocation").Value > 0 Then
                    If ugrow.Cells("Reason").Value = "" Then
                        If bolAnyReason Then
                            If MsgBox("Reallocation Reason Required.  Use:  '" & strLastReason & "' for this adjustment?", MsgBoxStyle.YesNo, "Missing Reason") = MsgBoxResult.Yes Then
                                ugrow.Cells("Reason").Value = strLastReason
                            Else
                                Return False
                            End If
                        Else
                            MsgBox("Reallocation Reason Required.", MsgBoxStyle.OKOnly, "Missing Reason")
                            Return False
                        End If
                    Else
                        bolAnyReason = True
                        strLastReason = ugrow.Cells("Reason").Value
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot ValidateData " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Function
End Class
