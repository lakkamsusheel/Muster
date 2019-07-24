Option Strict On
Option Explicit On 

' _____________________________________________________________________________________________'
'
' Invoice Request Administartor
' Created : Thomas Franey                   Ciber
' version 1.0
' Summary: Allows Finance Admin's to select Finance, and their Event Id to quickly add or update an 
'          invoice request

Public Class InvoiceRequestAdmin
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        StartPosition = FormStartPosition.Manual
        Location = New Point(10, 10)


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
    Friend WithEvents LblFacility As System.Windows.Forms.Label
    Friend WithEvents cboxFacility As System.Windows.Forms.ComboBox
    Friend WithEvents lblReimbursements As System.Windows.Forms.Label
    Friend WithEvents ugPayments As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cboxEvents As System.Windows.Forms.ComboBox
    Friend WithEvents mainControlPanel As System.Windows.Forms.Panel
    Friend WithEvents GoButton As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cboxFacility = New System.Windows.Forms.ComboBox
        Me.LblFacility = New System.Windows.Forms.Label
        Me.lblReimbursements = New System.Windows.Forms.Label
        Me.cboxEvents = New System.Windows.Forms.ComboBox
        Me.ugPayments = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.mainControlPanel = New System.Windows.Forms.Panel
        Me.GoButton = New System.Windows.Forms.Button
        CType(Me.ugPayments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.mainControlPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboxFacility
        '
        Me.cboxFacility.Location = New System.Drawing.Point(56, 8)
        Me.cboxFacility.Name = "cboxFacility"
        Me.cboxFacility.Size = New System.Drawing.Size(264, 21)
        Me.cboxFacility.TabIndex = 0
        '
        'LblFacility
        '
        Me.LblFacility.Location = New System.Drawing.Point(8, 8)
        Me.LblFacility.Name = "LblFacility"
        Me.LblFacility.Size = New System.Drawing.Size(48, 16)
        Me.LblFacility.TabIndex = 1
        Me.LblFacility.Text = "Facility :"
        Me.LblFacility.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblReimbursements
        '
        Me.lblReimbursements.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblReimbursements.Location = New System.Drawing.Point(368, 8)
        Me.lblReimbursements.Name = "lblReimbursements"
        Me.lblReimbursements.Size = New System.Drawing.Size(112, 16)
        Me.lblReimbursements.TabIndex = 3
        Me.lblReimbursements.Text = "Financial Events :"
        Me.lblReimbursements.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cboxEvents
        '
        Me.cboxEvents.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboxEvents.Enabled = False
        Me.cboxEvents.Location = New System.Drawing.Point(472, 8)
        Me.cboxEvents.Name = "cboxEvents"
        Me.cboxEvents.Size = New System.Drawing.Size(112, 21)
        Me.cboxEvents.TabIndex = 2
        Me.cboxEvents.Text = "ComboBox1"
        '
        'ugPayments
        '
        Me.ugPayments.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ugPayments.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPayments.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugPayments.Enabled = False
        Me.ugPayments.Location = New System.Drawing.Point(8, 40)
        Me.ugPayments.Name = "ugPayments"
        Me.ugPayments.Size = New System.Drawing.Size(576, 160)
        Me.ugPayments.TabIndex = 22
        '
        'mainControlPanel
        '
        Me.mainControlPanel.Controls.Add(Me.GoButton)
        Me.mainControlPanel.Controls.Add(Me.cboxFacility)
        Me.mainControlPanel.Controls.Add(Me.ugPayments)
        Me.mainControlPanel.Controls.Add(Me.LblFacility)
        Me.mainControlPanel.Controls.Add(Me.cboxEvents)
        Me.mainControlPanel.Controls.Add(Me.lblReimbursements)
        Me.mainControlPanel.Dock = System.Windows.Forms.DockStyle.Top
        Me.mainControlPanel.Location = New System.Drawing.Point(0, 0)
        Me.mainControlPanel.Name = "mainControlPanel"
        Me.mainControlPanel.Size = New System.Drawing.Size(592, 216)
        Me.mainControlPanel.TabIndex = 24
        '
        'GoButton
        '
        Me.GoButton.Location = New System.Drawing.Point(320, 8)
        Me.GoButton.Name = "GoButton"
        Me.GoButton.Size = New System.Drawing.Size(32, 24)
        Me.GoButton.TabIndex = 23
        Me.GoButton.Text = "Go"
        '
        'InvoiceRequestAdmin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 262)
        Me.Controls.Add(Me.mainControlPanel)
        Me.IsMdiContainer = True
        Me.Name = "InvoiceRequestAdmin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Invoice Request Administrator"
        CType(Me.ugPayments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.mainControlPanel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "private members"


    Private oFinancial As MUSTER.BusinessLogic.pFinancial
    Private ofacility As MUSTER.BusinessLogic.pFacility

    Private vFacilityID As Integer = -1
    Private vFinancialEventID As Integer = -1
    Private vReimbursementID As Integer = -1


#End Region





#Region "Private Properties"

    Private ReadOnly Property Financial() As MUSTER.BusinessLogic.pFinancial

        Get
            If oFinancial Is Nothing Then
                oFinancial = New MUSTER.BusinessLogic.pFinancial
            End If

            Return oFinancial
        End Get
    End Property



    Private ReadOnly Property Facility() As MUSTER.BusinessLogic.pFacility

        Get
            If ofacility Is Nothing Then
                ofacility = New MUSTER.BusinessLogic.pFacility
            End If

            Return ofacility
        End Get
    End Property

#End Region

#Region "Private Methods - Data Setting"

    Private Sub SetFinancialEvent(ByVal id As Integer)

        vFinancialEventID = id

        If vFinancialEventID > -2 Then

            Financial.Retrieve(vFinancialEventID)

        Else
            Financial.Clear()
        End If


        LoadPaymentsGrid()

    End Sub

    Private Sub SetReimbursement(ByVal id As Integer)

        If Not Me.ActiveMdiChild Is Nothing Then
            ActiveMdiChild.Close()
        End If

        vReimbursementID = id

        ShowInvoiceRequestManager()

    End Sub

    Private Sub SetFacility(ByVal id As Integer)

        vFacilityID = id
        If vFacilityID > -1 Then
            Facility.ID = vFacilityID
        Else
            Facility.Clear()
        End If

        LoadEventList()

    End Sub


#End Region

#Region "Private Methods -  Data Loading"

    Private Sub LoadEventList()

        If vFacilityID > -1 Then
            Dim dTemp As DataTable = Facility.FinancialEventDataset.Tables(0)

            With cboxEvents

                .DisplayMember = "FIN_EVENT_ID"
                .ValueMember = "FIN_EVENT_ID"
                .DataSource = dTemp
                cboxEvents.Enabled = True

            End With
        Else

            cboxEvents.Enabled = False
            ugPayments.DataSource = Nothing
            ugPayments.Enabled = False

        End If


    End Sub

    Private Sub ShowInvoiceRequestManager()

        Try
            Dim frmincompApplication As IncompleteApplication

            If IsNothing(frmincompApplication) Then

                frmincompApplication = New IncompleteApplication
            Else
                Exit Try
            End If


            With frmincompApplication

                .FinancialEventID = vFinancialEventID
                .FinancialReimbursementID = vReimbursementID
                .StartPosition = FormStartPosition.Manual
                .ControlBox = False

                .MdiParent = Me
                TopMost = True

            End With

            AddHandler frmincompApplication.Resize, AddressOf AdjustMDIMaster
            AddHandler frmincompApplication.Closed, AddressOf closeMdiChild

            mainControlPanel.Enabled = False


            frmincompApplication.Show()

            frmincompApplication.Location = New Point(0, 0)

            AddHandler frmincompApplication.Move, AddressOf MDIMoved

            LoadPaymentsGrid()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub LoadFacilityList()

        Dim dTemp As DataTable = Facility.GetTable("Select facility_id, (convert(nvarchar(100),facility_id) + ' ' + [name]) as Description from tblREG_FACILITY")

        With cboxFacility

            .DisplayMember = "Description"
            .ValueMember = "facility_id"
            .DataSource = dTemp
            .SelectedIndex = -1


        End With

    End Sub


    Private Sub LoadPaymentsGrid()
        Dim dsLocal As DataSet
        Dim tmpBand As Int16


        If vFinancialEventID > -2 Then

            dsLocal = oFinancial.PaymentGridDataset(False)
            'ugPayments.DataSource = dsLocal
            If dsLocal.Tables.Count > 0 Then
                dsLocal.Tables(0).DefaultView.Sort = "RECEIVED_DATE DESC"
                ugPayments.DataSource = dsLocal.Tables(0).DefaultView
            Else
                ugPayments.DataSource = dsLocal
            End If

            ugPayments.Rows.CollapseAll(True)
            ugPayments.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugPayments.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Payment Table have rows
                ugPayments.DisplayLayout.Bands(0).Columns("FIN_EVENT_ID").Hidden = True
                ugPayments.DisplayLayout.Bands(0).Columns("Reimbursement_ID").Hidden = True
                ugPayments.DisplayLayout.Bands(0).Columns("COMMITMENT_ID").Hidden = True
                ugPayments.DisplayLayout.Bands(0).Columns("Document_Location").Hidden = True
                ugPayments.DisplayLayout.Bands(0).Columns("rawRequestedAmount").Hidden = True
                ugPayments.DisplayLayout.Bands(0).Columns("rawPaidAmount").Hidden = True

                ugPayments.DisplayLayout.Bands(0).Columns("Received_date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugPayments.DisplayLayout.Bands(0).Columns("Received_date").Header.Caption = "Received"
                ugPayments.DisplayLayout.Bands(0).Columns("Received_date").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ugPayments.DisplayLayout.Bands(0).Columns("Requested_Amount").Header.Caption = "Requested"
                ugPayments.DisplayLayout.Bands(0).Columns("Requested_Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ugPayments.DisplayLayout.Bands(0).Columns("Requested_Invoiced").Header.Caption = "Requested(Inv)"
                ugPayments.DisplayLayout.Bands(0).Columns("Requested_Invoiced").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ugPayments.DisplayLayout.Bands(0).Columns("Paid").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ugPayments.DisplayLayout.Bands(0).Columns("Payment_Number").Header.Caption = "Pmt#"
                ugPayments.DisplayLayout.Bands(0).Columns("Payment_Number").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ugPayments.DisplayLayout.Bands(0).Columns("Payment_Date").Header.Caption = "Pmt Date"
                ugPayments.DisplayLayout.Bands(0).Columns("Payment_Date").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugPayments.DisplayLayout.Bands(0).Columns("Payment_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                ugPayments.DisplayLayout.Bands(0).Columns("ApprovalRequired").Header.Caption = "App Reqd"
                ugPayments.DisplayLayout.Bands(0).Columns("ApprovalRequired").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugPayments.DisplayLayout.Bands(0).Columns("Approved").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugPayments.DisplayLayout.Bands(0).Columns("On_Hold").Header.Caption = "On Hold"
                ugPayments.DisplayLayout.Bands(0).Columns("On_Hold").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugPayments.DisplayLayout.Bands(0).Columns("Incomplete").Header.Caption = "Incomplete"
                ugPayments.DisplayLayout.Bands(0).Columns("Incomplete").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


                ugPayments.DisplayLayout.Bands(0).Columns("Received_date").Width = 100
                ugPayments.DisplayLayout.Bands(0).Columns("Requested_Amount").Width = 100
                ugPayments.DisplayLayout.Bands(0).Columns("Requested_Invoiced").Width = 100
                ugPayments.DisplayLayout.Bands(0).Columns("Paid").Width = 100



            End If

            Me.ugPayments.Enabled = True
        Else

            Me.ugPayments.Enabled = False
            Me.ugPayments.DataSource = Nothing

        End If

    End Sub


#End Region




#Region "Form events"

    Sub AdjustMDIMaster(ByVal sender As Object, ByVal e As EventArgs)

        With DirectCast(sender, Form)

            Dim newHt As Integer = .Height + mainControlPanel.Height + 80

            MinimumSize = New Size(645, newHt)
            MaximumSize = New Size(1200, newHt)

        End With

    End Sub

    Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        LoadFacilityList()

        AddHandler cboxEvents.SelectedIndexChanged, AddressOf EventIDClick
        AddHandler cboxFacility.SelectedIndexChanged, AddressOf FacilityClick
        AddHandler ugPayments.DoubleClick, AddressOf ReImbursementDblClick


    End Sub

    Sub form_close(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Closed

        oFinancial = Nothing
        ofacility = Nothing
    End Sub


    Sub closeMdiChild(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Closed

        Me.mainControlPanel.Enabled = True

        Me.SetFinancialEvent(-2)

    End Sub

    Sub MDIMoved(ByVal sender As Object, ByVal e As EventArgs)

        With DirectCast(sender, Form)
            .Location = New Point(0, 0)
        End With


    End Sub


#End Region

#Region "ListClickEvents"

    Sub FacilityClick(ByVal sender As Object, ByVal e As EventArgs)

        If Not cboxFacility.SelectedValue Is Nothing Then
            SetFinancialEvent(-1)
            SetFacility(Convert.ToInt32(cboxFacility.SelectedValue))
        Else
            SetFacility(-1)
            SetFinancialEvent(-1)
        End If


    End Sub

    Sub EventIDClick(ByVal sender As Object, ByVal e As EventArgs)

        If Not cboxEvents.SelectedValue Is Nothing Then
            SetFinancialEvent(Convert.ToInt32(cboxEvents.SelectedValue))
        Else
            SetFinancialEvent(-3)
        End If
    End Sub

    Sub ReImbursementDblClick(ByVal sender As Object, ByVal e As EventArgs)
        With DirectCast(sender, Infragistics.Win.UltraWinGrid.UltraGrid)

            If Not .ActiveRow Is Nothing Then
                SetReimbursement(Convert.ToInt32(.ActiveRow.Cells("Reimbursement_ID").Value))

            End If

        End With
    End Sub


    Private Sub GoButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoButton.Click

        With cboxFacility

            cboxFacility.SelectedIndex = -1

            If IsNumeric(.Text) Then
                cboxFacility.SelectedValue = Convert.ToInt32(.Text)

            ElseIf .Text.IndexOf(" ") <> -1 AndAlso IsNumeric(.Text.Substring(0, .Text.IndexOf(" "))) Then
                cboxFacility.SelectedValue = Convert.ToInt32(.Text.Substring(0, .Text.IndexOf(" ")))

            Else
                .SelectedIndex = -1

            End If


        End With


    End Sub

#End Region



End Class
